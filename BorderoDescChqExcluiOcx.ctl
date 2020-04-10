VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoDescChqExcluiOcx 
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   6615
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1275
      Left            =   180
      TabIndex        =   11
      Top             =   765
      Width           =   6315
      Begin VB.Label ContaCorrente 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4395
         TabIndex        =   19
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label DataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1260
         TabIndex        =   18
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label CarteiraCobranca 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4395
         TabIndex        =   17
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Cobrador 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1260
         TabIndex        =   16
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Index           =   4
         Left            =   435
         TabIndex        =   14
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
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
         Index           =   0
         Left            =   3585
         TabIndex        =   13
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
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
         Left            =   3015
         TabIndex        =   12
         Top             =   885
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cheques"
      Height          =   810
      Left            =   180
      TabIndex        =   6
      Top             =   2145
      Width           =   6315
      Begin VB.Label ValorCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4485
         TabIndex        =   10
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label Label5 
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
         Height          =   210
         Left            =   3900
         TabIndex        =   9
         Top             =   345
         Width           =   525
      End
      Begin VB.Label QtdeCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   705
         TabIndex        =   7
         Top             =   345
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4815
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   1680
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BorderoDescChqExcluiOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "BorderoDescChqExcluiOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BorderoDescChqExcluiOcx.ctx":02D8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoTrazer 
      Caption         =   "Trazer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3735
      TabIndex        =   2
      Top             =   240
      Width           =   795
   End
   Begin MSMask.MaskEdBox NumBordero 
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   247
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelNumBordero 
      Caption         =   "Número do Borderô:"
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
      Left            =   375
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   300
      Width           =   1740
   End
End
Attribute VB_Name = "BorderoDescChqExcluiOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoBorderoDescChq As AdmEvento
Attribute objEventoBorderoDescChq.VB_VarHelpID = -1

Public iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Set objEventoBorderoDescChq = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Exit Sub
    
Erro_Form_Unload:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143738)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    Set objEventoBorderoDescChq = New AdmEvento

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143739)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoDescChq As ClassBorderoDescChq) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se objBorderoDescChq está instanciado
    If Not objBorderoDescChq Is Nothing Then
    
        'tenta trazê-lo para a tela
        lErro = Traz_BorderoDescChq_Tela(objBorderoDescChq)
        If lErro <> SUCESSO And lErro <> 109290 Then gError 109289
        
        'se não encontrou
        If lErro = 109290 Then
            
            'limpa a tela
            Call Limpa_Tela_BorderoDescChq
            
            'e preenche o código do borderô
            NumBordero.Text = objBorderoDescChq.lNumBordero
            
        End If
            
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 109289
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143740)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objBorderoDescChq As New ClassBorderoDescChq
Dim vbResp As VbMsgBoxResult

On Error GoTo Erro_BotaoGravar_Click
    
    'transforma o mouse em apulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o número do borderô não estiver preenchido-> erro
    If Len(Trim(NumBordero.Text)) = 0 Then gError 109305
    
    'preenche os dados do borderô
    Call Move_Tela_Memoria(objBorderoDescChq)
    
    'busca o borderô
    lErro = CF("BorderoDescChq_Le", objBorderoDescChq)
    If lErro <> SUCESSO And lErro <> 109291 Then gError 109338
    
    'se não achou-> erro
    If lErro = 109291 Then gError 109339
    
    'pergunta se tem certeza
    vbResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_BORDERODESCCHQ", objBorderoDescChq.iFilialEmpresa, objBorderoDescChq.lNumBordero)
    
    'se responder que sim
    If vbResp = vbYes Then
    
        'exclui o bordero
        lErro = CF("BorderoDescChq_Exclui", objBorderoDescChq)
        If lErro <> SUCESSO Then gError 109306
    
        'limpa a tela
        Call Limpa_Tela_BorderoDescChq
    
    End If
    
    'volta o mouse para o estado normal
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
        
        Case 109305
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", gErr)
            
        Case 109306, 109308
        
        Case 109339
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERODESCCHQ_NAO_ENCONTRADO", gErr, objBorderoDescChq.iFilialEmpresa, objBorderoDescChq.lNumBordero)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143741)
    
    End Select
    
    'volta o mouse para o estado normal
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
    Call Limpa_Tela_BorderoDescChq
    
    iAlterado = 0

End Sub

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim objBorderoDescChq As New ClassBorderoDescChq

On Error GoTo Erro_BotaoTrazer_Click

    'se o número do borderô não estiver preenchido->erro
    If Len(Trim(NumBordero.Text)) = 0 Then gError 109302
    
    'preenche os dados de um borderô para a busca
    Call Move_Tela_Memoria(objBorderoDescChq)
    
    'preenche a tela
    lErro = Traz_BorderoDescChq_Tela(objBorderoDescChq)
    If lErro <> SUCESSO And lErro <> 109290 Then gError 109303
    
    'se não encontrou-> erro
    If lErro = 109290 Then gError 109304

    Exit Sub
    
Erro_BotaoTrazer_Click:
    
    Select Case gErr
    
        Case 109302
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", gErr)
    
        Case 109303
        
        Case 109304
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERODESCCHQ_NAO_ENCONTRADO", gErr, objBorderoDescChq.iFilialEmpresa, objBorderoDescChq.lNumBordero)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143742)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelNumBordero_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objBorderoDescChq As New ClassBorderoDescChq

On Error GoTo Erro_LabelNumBordero_Click

    'se o número do borderô estiver preenchido
    If Len(Trim(NumBordero.Text)) <> 0 Then Call Move_Tela_Memoria(objBorderoDescChq)
    
    'chama o browser
    Call Chama_Tela("BorderoDescChqLista", colSelecao, objBorderoDescChq, objEventoBorderoDescChq)

    Exit Sub
    
Erro_LabelNumBordero_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143743)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NumBordero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumBordero_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumBordero, iAlterado)
End Sub

Private Sub NumBordero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumBordero_Validate

    'se o número do borderô estiver preenchido
    If Len(Trim(NumBordero.Text)) <> 0 Then
    
        'critica
        lErro = Long_Critica(NumBordero.Text)
        If lErro <> SUCESSO Then gError 109301
    
    End If
    
    Exit Sub
    
Erro_NumBordero_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 109301
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143744)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoBorderoDescChq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objBorderoDescChq As ClassBorderoDescChq

On Error GoTo Erro_objEventoBorderoDescChq_evSelecao

    'aponta para o obj recebido por parâmetro
    Set objBorderoDescChq = obj1
    
    lErro = Traz_BorderoDescChq_Tela(objBorderoDescChq)
    If lErro <> SUCESSO And lErro <> 109290 Then gError 109338
    
    If lErro = 109290 Then gError 109339
    
'    lErro = CF("BorderoDescChq_Le", objBorderoDescChq)
'    If lErro <> SUCESSO And lErro <> 109291 Then gError 109297
'
'    lErro = CF("BorderoDescChq_Le_ChequesPre", objBorderoDescChq.lNumBordero, objBorderoDescChq.colchequepre)
'    if lerro
'
'
'    'se não encontrou-> erro
'    If lErro = 109291 Then gError 109298
'
'    'preenche a tela
'    NumBordero.Text = objBorderoDescChq.lNumBordero
'    Cobrador.Caption = objBorderoDescChq.iCobrador & SEPARADOR & objBorderoDescChq.sCobrador
'    CarteiraCobranca.Caption = objBorderoDescChq.iCarteiraCobranca & SEPARADOR & objBorderoDescChq.sCarteiraCobranca
'    ContaCorrente.Caption = objBorderoDescChq.iContaCorrente & SEPARADOR & objBorderoDescChq.sContaCorrente
'    DataEmissao.Caption = Format(objBorderoDescChq.dtDataEmissao, "dd/mm/yyyy")
'    QtdeCheques.Caption = objBorderoDescChq.iQuantChequesSel
'    ValorCheques.Caption = Format(objBorderoDescChq.dValorChequesSel, "STANDARD")
'
    Me.Show
    
    Exit Sub
    
Erro_objEventoBorderoDescChq_evSelecao:
    
    Select Case gErr
    
        Case 109338
        
        Case 109339
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERODESCCHQ_NAO_ENCONTRADO", gErr, objBorderoDescChq.iFilialEmpresa, objBorderoDescChq.lNumBordero)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143745)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objBorderoDescChq As New ClassBorderoDescChq

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderoDescChq"

    'Armazena os dados presentes na tela em objOperador
    Call Move_Tela_Memoria(objBorderoDescChq)

    'preenche a coleção de campos valores
    colCampoValor.Add "FilialEmpresa", objBorderoDescChq.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumBordero", objBorderoDescChq.lNumBordero, 0, "NumBordero"
    colCampoValor.Add "ContaCorrente", objBorderoDescChq.iContaCorrente, 0, "ContaCorrente"
    colCampoValor.Add "DataEmissao", objBorderoDescChq.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataContabil", objBorderoDescChq.dtDataContabil, 0, "DataContabil"
    colCampoValor.Add "Cobrador", objBorderoDescChq.iCobrador, 0, "Cobrador"
    colCampoValor.Add "CarteiraCobranca", objBorderoDescChq.iCarteiraCobranca, 0, "CarteiraCobranca"
    colCampoValor.Add "DataDeposito", objBorderoDescChq.dtDataDeposito, 0, "DataDeposito"
    colCampoValor.Add "ValorCredito", objBorderoDescChq.dvalorCredito, 0, "ValorCredito"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143746)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objBorderoDescChq As New ClassBorderoDescChq
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'preenche os atributos de objBorderoDescChq
    objBorderoDescChq.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objBorderoDescChq.lNumBordero = colCampoValor.Item("NumBordero").vValor
    objBorderoDescChq.iContaCorrente = colCampoValor.Item("ContaCorrente").vValor
    objBorderoDescChq.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
    objBorderoDescChq.dtDataContabil = colCampoValor.Item("DataContabil").vValor
    objBorderoDescChq.iCobrador = colCampoValor.Item("Cobrador").vValor
    objBorderoDescChq.iCarteiraCobranca = colCampoValor.Item("CarteiraCobranca").vValor
    objBorderoDescChq.dtDataDeposito = colCampoValor.Item("DataDeposito").vValor
    objBorderoDescChq.dvalorCredito = colCampoValor.Item("ValorCredito").vValor

    'preenche a tel
    lErro = Traz_BorderoDescChq_Tela(objBorderoDescChq)
    If lErro <> SUCESSO Then gError 109299
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 109299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143747)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoDescChq()

    Call Limpa_Tela(Me)
    
    'limpa cada label
    Cobrador.Caption = ""
    CarteiraCobranca.Caption = ""
    ContaCorrente.Caption = ""
    DataEmissao.Caption = ""
    QtdeCheques.Caption = ""
    ValorCheques.Caption = ""
    
    iAlterado = 0

End Sub

Private Function Traz_BorderoDescChq_Tela(ByVal objBorderoDescChq As ClassBorderoDescChq) As Long

Dim lErro As Long
Dim dValorChequesSel As Double
Dim objChequePre As ClassChequePre

On Error GoTo Erro_Traz_BorderoDescChq_Tela

    'lê o borderoDescChq
    lErro = CF("BorderoDescChq_Le", objBorderoDescChq)
    If lErro <> SUCESSO And lErro <> 109291 Then gError 109292
    
    'se não encontrou-> erro
    If lErro = 109291 Then gError 109290
    
    'lê os cheques vinculados ao borderodescChq em questão
    lErro = CF("BorderoDescChq_Le_ChequesPre", objBorderoDescChq.lNumBordero, objBorderoDescChq.colchequepre)
    If lErro <> SUCESSO And lErro <> 109333 Then gError 109334
    
    'se não encontrou nenhum cheque-> erro
    If lErro = 109333 Then gError 109335
    
    'preenche a tela
    NumBordero.Text = objBorderoDescChq.lNumBordero
    Cobrador.Caption = objBorderoDescChq.iCobrador & SEPARADOR & objBorderoDescChq.sCobrador
    CarteiraCobranca.Caption = objBorderoDescChq.iCarteiraCobranca & SEPARADOR & objBorderoDescChq.sCarteiraCobranca
    ContaCorrente.Caption = objBorderoDescChq.iContaCorrente & SEPARADOR & objBorderoDescChq.sContaCorrente
    DataEmissao.Caption = Format(objBorderoDescChq.dtDataEmissao, "dd/mm/yyyy")
    QtdeCheques.Caption = objBorderoDescChq.iQuantChequesSel
    ValorCheques.Caption = Format(objBorderoDescChq.dValorChequesSel, "STANDARD")
    
    'totaliza os cheques associados ao borderô
    For Each objChequePre In objBorderoDescChq.colchequepre
        dValorChequesSel = dValorChequesSel + objChequePre.dValor
    Next
    
    'aproveita e preenche os totalizadores do objeto
    objBorderoDescChq.dValorChequesSel = dValorChequesSel
    objBorderoDescChq.iQuantChequesSel = objBorderoDescChq.colchequepre.Count
    
    'preenche os totais da tela
    ValorCheques.Caption = Format(objBorderoDescChq.dValorChequesSel, "STANDARD")
    QtdeCheques.Caption = objBorderoDescChq.iQuantChequesSel
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 109293

    Traz_BorderoDescChq_Tela = SUCESSO
    
    iAlterado = 0
    
    Exit Function
    
Erro_Traz_BorderoDescChq_Tela:
    
    Traz_BorderoDescChq_Tela = gErr
    
    Select Case gErr
    
        Case 109290, 109292, 109293, 109334, 109335
        '109290 tratado na chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143748)
    
    End Select
    
    Exit Function

End Function



Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Move_Tela_Memoria(objBorderoDescChq As ClassBorderoDescChq)

    objBorderoDescChq.iCobrador = Codigo_Extrai(Cobrador.Caption)
    objBorderoDescChq.iContaCorrente = Codigo_Extrai(ContaCorrente.Caption)
    objBorderoDescChq.iCarteiraCobranca = Codigo_Extrai(CarteiraCobranca.Caption)
    objBorderoDescChq.dtDataEmissao = StrParaDate(DataEmissao.Caption)
    objBorderoDescChq.iQuantChequesSel = StrParaInt(QtdeCheques.Caption)
    objBorderoDescChq.dValorChequesSel = StrParaDbl(ValorCheques.Caption)
    objBorderoDescChq.lNumBordero = StrParaLong(NumBordero.Text)
    objBorderoDescChq.iFilialEmpresa = giFilialEmpresa

End Sub

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

On Error GoTo Erro_GeraContabilizacao
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143749)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

On Error GoTo Erro_Calcula_Mnemonico
    
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143750)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_BORDERO_PAGT_P4
    Set Form_Load_Ocx = Me
    Caption = "Excluir Borderô de Desconto de Cheques"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoDescChqExclui"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then Call LabelNumBordero_Click
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

'***** fim do trecho a ser copiado ******
