VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTChequePagAvulso3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjChequesPagAvulso As ClassChequesPagAvulso 'Alterado por Wagner
Dim iChequeImpresso As Integer

Function Trata_Parametros(Optional objChequesPagAvulso As ClassChequesPagAvulso) As Long

Dim lErro As Long
Dim objCtaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Trata_Parametros

    Set gobjChequesPagAvulso = objChequesPagAvulso
    
    objCtaCorrenteInt.iCodigo = gobjChequesPagAvulso.iCta
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objCtaCorrenteInt.iCodigo, objCtaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 59085
    
    LabelConta.Caption = CStr(objCtaCorrenteInt.iCodigo) & SEPARADOR & objCtaCorrenteInt.sNomeReduzido
    LabelValorCheque.Caption = Format(gobjChequesPagAvulso.objInfoChequePag.dValor, "Standard")
    Nominal.Text = gobjChequesPagAvulso.objInfoChequePag.sFavorecido 'Alterado por Wagner
    If gobjChequesPagAvulso.lNumCheque <> 0 Then NumCheque.Text = CStr(gobjChequesPagAvulso.lNumCheque)
    
    '########################################
    'Inserido por Wagner
    lErro = CF2(Me, "CHQA_Trata_Parametros", gobjChequesPagAvulso)
    If lErro <> SUCESSO Then gError 132569
    '########################################
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 59085, 132569 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155514)

    End Select

    Exit Function

End Function

'--- inicio dos properties get dos controles da tela

Public Property Get Nominal() As Object
     Set Nominal = objUserControl.Controls("Nominal")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get BotaoVoltar() As Object
     Set BotaoVoltar = objUserControl.Controls("BotaoVoltar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoSeguir() As Object
     Set BotaoSeguir = objUserControl.Controls("BotaoSeguir")
End Property

Public Property Get ConfigurarImpressao() As Object
     Set ConfigurarImpressao = objUserControl.Controls("ConfigurarImpressao")
End Property

Public Property Get NumCheque() As Object
     Set NumCheque = objUserControl.Controls("NumCheque")
End Property

Public Property Get Imprimir() As Object
     Set Imprimir = objUserControl.Controls("Imprimir")
End Property

Public Property Get LabelConta() As Object
     Set LabelConta = objUserControl.Controls("LabelConta")
End Property

Public Property Get LabelValorCheque() As Object
     Set LabelValorCheque = objUserControl.Controls("LabelValorCheque")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoSeguir_Click()

Dim vbMsgRes As VbMsgBoxResult, lErro As Long
    
On Error GoTo Erro_BotaoSeguir_Click

    If iChequeImpresso = 0 Then
    
        'perguntar se prossegue mesmo sem ter impresso
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NAO_IMP_CHEQUE")
        If vbMsgRes <> vbYes Then Exit Sub
        
    End If
    
    'Verifica se o NumCheque está preenchido
    If Len(Trim(NumCheque.Text)) = 0 Then Error 59086

    gobjChequesPagAvulso.lNumCheque = CLng(NumCheque.Text)
    gobjChequesPagAvulso.objInfoChequePag.lNumRealCheque = gobjChequesPagAvulso.lNumCheque

    '##########################
    'Inserido por Wagner
    gobjChequesPagAvulso.objInfoChequePag.sFavorecido = Nominal.Text
    gobjChequesPagAvulso.objInfoChequePag.sObservacao = Observacao.Text
    '##########################

    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequePagAvulso4", gobjChequesPagAvulso)
       
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoSeguir_Click:

    Select Case Err
          
        Case 59086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXCHEQUE_NAO_INFORMADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155515)
     
    End Select
     
    Exit Sub
    
End Sub

Public Sub BotaoVoltar_Click()

    'Chama a tela do passo Anterior
    Call Chama_Tela("ChequePagAvulso2", gobjChequesPagAvulso)
       
    'Fecha a tela
    Unload Me
    
End Sub

Public Sub ConfigurarImpressao_Click()

    Call Sist_ImpressoraDlg(1)

End Sub

Public Sub Form_Load()

    iChequeImpresso = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjChequesPagAvulso = Nothing
    
End Sub

Public Sub Imprimir_Click()

Dim lErro As Long

On Error GoTo Erro_ImprimirTeste_Click

    'Verifica se o NumCheque está preenchido
    If Len(Trim(NumCheque.Text)) = 0 Then Error 15700

    '##########################
    'Inserido por Wagner
    gobjChequesPagAvulso.objInfoChequePag.sFavorecido = Nominal.Text
    gobjChequesPagAvulso.objInfoChequePag.sObservacao = Observacao.Text
    '##########################

    lErro = CF("ChequesPagAvulso_PrepararImpressao", gobjChequesPagAvulso)
    If lErro <> SUCESSO Then Error 7791

    lErro = ImprimirCheques(gobjChequesPagAvulso.lNumImpressao, gobjChequesPagAvulso.sLayoutCheque, IIf(gobjChequesPagAvulso.dtBomPara <> DATA_NULA, gobjChequesPagAvulso.dtBomPara, gobjChequesPagAvulso.dtEmissao))
    If lErro <> SUCESSO Then Error 7792

    iChequeImpresso = 1
    
    Exit Sub
    
Erro_ImprimirTeste_Click:

    Select Case Err

        Case 7791, 7792
        
        Case 15700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXCHEQUE_NAO_INFORMADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155516)

    End Select

    Exit Sub
    
End Sub

Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19466
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade, "TIGNORARMARGEM", "S")
    If lErro <> SUCESSO Then Error 7431

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = Err

    Select Case Err

        Case 7431, 19466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155517)

    End Select

    Exit Function

End Function

Public Sub NumCheque_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumCheque)

End Sub

Public Sub NumCheque_Validate(bCancel As Boolean)

Dim lErro As Long, objMovCCI As New ClassMovContaCorrente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NumCheque_Validate

    'Verifica se o NumCheque está preenchido
    If Len(Trim(NumCheque.Text)) = 0 Then Exit Sub

    'Verifica se o NumCheque que está na tela é um número do tipo Long
    lErro = Long_Critica(NumCheque.Text)
    If lErro <> SUCESSO Then Error 43058
    
    objMovCCI.iCodConta = gobjChequesPagAvulso.iCta
    objMovCCI.lNumero = CLng(NumCheque.Text)
    lErro = CF("MovCCI_Pesquisa_Cheque", objMovCCI)
    If lErro <> SUCESSO And lErro <> 59082 Then Error 59083
    If lErro <> SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CHEQUE_NUM_USADO_DATA", objMovCCI.dtDataMovimento)
        If vbMsgRes = vbYes Then Error 59084
    
    End If

    Exit Sub

Erro_NumCheque_Validate:

    Select Case Err

        Case 43058, 59083, 59084
            bCancel = True
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155518)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUE_MANUAL_P3
    Set Form_Load_Ocx = Me
    Caption = "Cheque Manual - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequePagAvulso3"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

'Private Sub LabelEmNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelEmNomeDe, Source, X, Y)
'End Sub
'
'Private Sub LabelEmNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelEmNomeDe, Button, Shift, X, Y)
'End Sub

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

'###############################################
'Inserido por Wagner
Public Function CHQA_Trata_Parametros(ByVal objCT As Object, ByVal objChequesPagAvulso As ClassChequesPagAvulso) As Long
    CHQA_Trata_Parametros = SUCESSO
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
    
    End If
    
End Sub
'###############################################

