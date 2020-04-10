VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ImportLctosOcx 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   ScaleHeight     =   1560
   ScaleWidth      =   6270
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   1950
      Picture         =   "ImportLctosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   825
      Width           =   990
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   3300
      Picture         =   "ImportLctosOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   810
      Width           =   990
   End
   Begin VB.TextBox Arquivo 
      Height          =   300
      Left            =   975
      TabIndex        =   1
      Top             =   285
      Width           =   4005
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "Procurar..."
      Height          =   540
      Left            =   5085
      Picture         =   "ImportLctosOcx.ctx":025C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2730
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
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
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   315
      Width           =   690
   End
End
Attribute VB_Name = "ImportLctosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoOK_Click()
Dim lErro As Long
Dim dtData As Date, lQtdeLctos As Long

On Error GoTo Erro_BotaoOK_Click

    'Verifica se TextBox Arquivo foi preenchido. Se nao foi, erro
    If Arquivo.Text = "" Then gError 22022

    lErro = ImportCtb_ObtemInfoArq(Arquivo.Text, giFilialEmpresa, dtData)
    If lErro <> SUCESSO Then gError 22021

    If Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_IMPCTB", dtData) = vbYes Then
    
        lErro = ImportCtb_ImportarArq(Arquivo.Text, giFilialEmpresa)
        If lErro <> SUCESSO Then gError 22021
    
    End If
    
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 22021

        Case 22022
             Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", Err)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'poderia pegar o ultimo diretorio
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        'Case 22019

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    '
    
End Sub

Private Sub BotaoCancelar_Click()

    Unload Me

End Sub

Private Sub BotaoProcurar_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo Erro_BotaoProcurar_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Arquivo.Text = CommonDialog1.FileName
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_EXTRATO_BANCARIO_CNAB
    Set Form_Load_Ocx = Me
    Caption = "Importação de Lançamentos Contábeis"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ImportLctos"
    
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
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Function ImportCtb_ObtemInfoArq(ByVal sNomeArq As String, ByVal iFilialEmpresa As Integer, dtData As Date) As Long
'abre o arquivo a ser importado e confere o cnpj e obtem a sua data de geracao
'tb verifica se este arquivo já foi importado antes

Dim lErro As Long, bArqAberto As Boolean
Dim fso As New FileSystemObject, ts As TextStream
Dim objFilial As New AdmFiliais, sRegistro As String
Dim bValidaEmp As Boolean

On Error GoTo Erro_ImportCtb_ObtemInfoArq

    bValidaEmp = True

    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(sNomeArq, 1, 0)
    bArqAberto = True

    'Até chegar ao fim do arquivo
    If ts.AtEndOfLine Then gError 184468
    
    'Busca o próximo registro do arquivo
    sRegistro = ts.ReadLine
                                
    If left(sRegistro, 2) <> "01" Then gError 184469
    
    objFilial.iCodFilial = iFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ImportCtb_Valida_Emp", sNomeArq, bValidaEmp)
    'lErro = ImportCtb_Valida_Emp(sNomeArq, bValidaEmp)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'If objFilial.sCgc <> Mid(sRegistro, 3, 14) And InStr(UCase(sNomeArq), "SP") = 0 Then gError 184470 'o arquivo de sp serve para as filiais PR e av brasil
    If objFilial.sCgc <> Mid(sRegistro, 3, 14) And bValidaEmp Then gError 184470
    
    dtData = StrParaDate(Mid(sRegistro, 17, 2) & "/" & Mid(sRegistro, 19, 2) & "/" & Mid(sRegistro, 21, 4))
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False

    ImportCtb_ObtemInfoArq = SUCESSO
    
    Exit Function
    
Erro_ImportCtb_ObtemInfoArq:

    ImportCtb_ObtemInfoArq = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 184468
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTB_ARQVAZIO", gErr)
            
        Case 184469
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTB_ARQ_INVALIDO", gErr)
        
        Case 184470
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTB_CNPJ_FILIAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Private Function ImportCtb_ImportarArq(ByVal sNomeArq As String, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, dtDataImportacao As Date, objContaCcl As New ClassContaCcl, dtData As Date
Dim lTransacao As Long, alComando(1 To 10) As Long, iIndice As Integer, lNumIntArq As Long
Dim sConta As String, sCcl As String, dValor As Double, sHistorico As String, sDocOrigem As String
Dim objLancamento_Detalhe As ClassLancamento_Detalhe, colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho, iSeq As Integer, dtDataAnterior As Date, iSeqArq As Integer

On Error GoTo Erro_ImportCtb_ImportarArq
        
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184471
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184472
    
    'verificar se o arquivo já foi importado
    lErro = Comando_Executar(alComando(1), "SELECT DataImportacao FROM ImportCtbArq WHERE FilialEmpresa = ? AND NomeArquivo = ?", _
        dtDataImportacao, iFilialEmpresa, sNomeArq)
    If lErro <> AD_SQL_SUCESSO Then gError 184473
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184474
    
    'se o arquivo já foi importado antes
    If lErro = AD_SQL_SUCESSO Then gError 184475
    
    'joga do arquivo texto para o banco de dados
    lErro = ImportCtb_ImportarArq1(sNomeArq, iFilialEmpresa, alComando(2), lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)
    sHistorico = String(STRING_HISTORICO, 0)
    sDocOrigem = String(STRING_DOCORIGEM, 0)
    lErro = Comando_Executar(alComando(3), "SELECT Seq, Data, Conta, Ccl, Historico, Valor, DocOrigem FROM ImportCtbLctos WHERE NumIntArq = ? ORDER BY Data, Seq", _
        iSeqArq, dtData, sConta, sCcl, sHistorico, dValor, sDocOrigem, lNumIntArq)
    If lErro <> AD_SQL_SUCESSO Then gError 184477
        
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184478
    
    dtDataAnterior = DATA_NULA
    
    Do While lErro = AD_SQL_SUCESSO
    
        If dtData <> dtDataAnterior Then
        
            If dtDataAnterior <> DATA_NULA Then
            
                lErro = ImportCtb_GravaDoc(objLancamento_Cabecalho, colLancamento_Detalhe)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            End If
            
            Set colLancamento_Detalhe = New Collection
            Set objLancamento_Cabecalho = New ClassLancamento_Cabecalho

            objLancamento_Cabecalho.dtData = dtData
            objLancamento_Cabecalho.sOrigem = "FLH"
            objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
            objLancamento_Cabecalho.lNumIntDoc = lNumIntArq
            
            iSeq = 0
            
            dtDataAnterior = dtData
            
        End If
        
        'adiciona registro em colLancamento_Detalhe
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
    
        iSeq = iSeq + 1
        
        'If left(sConta, 1) = "1" Or left(sConta, 1) = "2" Then sCcl = ""
        
        lErro = CF("ImportCtb_Trata_Conta_Ccl", sConta, sCcl)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        With objLancamento_Detalhe
            .lNumIntDoc = lNumIntArq
            .sDocOrigem = sDocOrigem
            .sConta = sConta
            .sCcl = sCcl
            .iSeq = iSeq
            .sHistorico = sHistorico
            .sProduto = ""
            .sOrigem = objLancamento_Cabecalho.sOrigem
            .dValor = dValor
            .sDocOrigem = sDocOrigem
        End With
                
        If Len(Trim(sCcl)) <> 0 Then
        
            Set objContaCcl = New ClassContaCcl
            objContaCcl.sConta = sConta
            objContaCcl.sCcl = sCcl
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then gError 184700
        
        End If
        
        colLancamento_Detalhe.Add objLancamento_Detalhe
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184481
    
    Loop
    
    If colLancamento_Detalhe.Count <> 0 Then
    
        lErro = ImportCtb_GravaDoc(objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184483
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ImportCtb_ImportarArq = SUCESSO
    
    Exit Function
    
Erro_ImportCtb_ImportarArq:

    ImportCtb_ImportarArq = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 184471
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 184472
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 184473, 184474
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCTBARQ", gErr)

        Case 184475
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTBARQ_JA_IMPORTADO", gErr)
            
        Case 184477, 184478, 184481
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCTBLCTOS", gErr)

        Case 184700
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", gErr, objContaCcl.sConta, objContaCcl.sCcl)
            
        Case 184483
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184484)

    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ImportCtb_GravaDoc(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
'obter lote e doc

Dim lErro As Long, objPeriodo As New ClassPeriodo, objContabAutoAux As New ClassContabAutoAux
Dim objLote As New ClassLote, lDoc As Long, iTotaisIguais As Integer
Dim alComando(1 To 2) As Long, iIndice As Integer

On Error GoTo Erro_ImportCtb_GravaDoc

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError ERRO_SEM_MENSAGEM
    Next
    
    lErro = CF("Periodo_Le1", objLancamento_Cabecalho.dtData, objPeriodo, objLancamento_Cabecalho.iFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.iPeriodoLote = objPeriodo.iPeriodo
    
    'obter numero do lote
    objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo
    objLote.sOrigem = objLancamento_Cabecalho.sOrigem
    lErro = CF("Lote_Automatico1", objLote)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'abrir o lote
    objLote.iStatus = LOTE_DESATUALIZADO
    lErro = CF("LotePendente_Grava_Trans", objLote)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Voucher_Automatico_Obter_Batch", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc, alComando(1))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objContabAutoAux.lProxDoc = lDoc + 1
    'atualiza o numero do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE ExercicioOrigem SET Doc = ?", alComando(1), objContabAutoAux.lProxDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184488
            
    objLancamento_Cabecalho.iLote = objLote.iLote
    objLancamento_Cabecalho.lDoc = lDoc
    
    lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'obtem os valores de debitos, creditos e a qtde de lancamentos lendo no bd
    lErro = CF("LanPendente_Critica_TotaisLote", objLote, iTotaisIguais)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'copia os valores lidos do bd p/os valores "informados"
    objLote.dTotInf = objLote.dTotDeb
    objLote.iNumDocInf = objLote.iNumDocAtual
    objLote.iNumLancInf = objLote.iNumLancAtual
    
    lErro = CF("LotePendente_Grava_Totais_Auto", objLote)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ImportCtb_GravaDoc = SUCESSO
    
    Exit Function
    
Erro_ImportCtb_GravaDoc:

    ImportCtb_GravaDoc = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 184488 'ERRO_ATUALIZACAO_EXERCICIOORIGEM_BATCH
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM_BATCH", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184489)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ImportCtb_ImportarArq1(ByVal sNomeArq As String, ByVal iFilialEmpresa As Integer, lComando As Long, lNumIntArq As Long) As Long

Dim lErro As Long, bArqAberto As Boolean ', bArqSP As Boolean
Dim fso As New FileSystemObject, ts As TextStream, iAux As Integer
Dim objFilial As New AdmFiliais, dSomaCreditos As Double, lQtde As Long
Dim iTipoReg As Integer, sRegistro As String, dtData As Date
Dim sConta As String, sConta1 As String, sCcl As String, sCcl1 As String, dValor As Double, sHistorico As String, sDocOrigem As String
Dim bValidaEmp As Boolean, iContaPreenchida As Integer, iCclPreenchida As Integer, bImportaLcto As Boolean
Dim objCcl As ClassCcl

On Error GoTo Erro_ImportCtb_ImportarArq1
    
    'bArqSP = (InStr(UCase(sNomeArq), "SP") <> 0)
    bValidaEmp = True
    bImportaLcto = True
    
    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_IMPORTCTB", lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(sNomeArq, 1, 0)
    bArqAberto = True
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        If Len(Trim(sRegistro)) <> 0 Then
            
            iTipoReg = StrParaInt(left(sRegistro, 2))
            
            Select Case iTipoReg
            
                Case 1
                    objFilial.iCodFilial = iFilialEmpresa
                    lErro = CF("FilialEmpresa_Le", objFilial)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    lErro = CF("ImportCtb_Valida_Emp", sNomeArq, bValidaEmp)
                    'lErro = ImportCtb_Valida_Emp(sNomeArq, bValidaEmp)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    'If objFilial.sCgc <> Mid(sRegistro, 3, 14) And (Not bArqSP) Then Error 184494
                    If objFilial.sCgc <> Mid(sRegistro, 3, 14) And bValidaEmp Then gError 184494 'Filial diferente da esperada
            
                Case 2
                
                    lQtde = lQtde + 1
                    dtData = StrParaDate(Mid(sRegistro, 3, 2) & "/" & Mid(sRegistro, 5, 2) & "/" & Mid(sRegistro, 7, 4))
                    sConta1 = Trim(Mid(sRegistro, 11, 20))
'                    sConta = ""
'                    For iAux = 1 To Len(sConta1)
'                        If Mid(sConta1, iAux, 1) <> "." Then sConta = sConta & Mid(sConta1, iAux, 1)
'                    Next
                    'If Len(sConta) < 11 Then sConta = sConta & String(11 - Len(sConta), "0")
                    
                    lErro = Conta_Formata_Importacao(sConta1, sConta, iContaPreenchida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If Len(Trim(sConta)) = 0 Then gError 184497
                    
                    lErro = CF("PlanoConta_Le_Conta", sConta)
                    If lErro <> SUCESSO And lErro <> 10051 Then gError ERRO_SEM_MENSAGEM
                    If lErro <> SUCESSO Then gError 184498
                    
                    If InStr(1, UCase(gsNomeEmpresa), "INPAL") <> 0 Then
                        sCcl1 = Trim(Mid(sRegistro, 36, 5))
                        If Len(sCcl1) = 4 Then sCcl1 = sCcl1 & "0"
                    Else
                        sCcl1 = Trim(Mid(sRegistro, 31, 10))
                    End If
                    
                    lErro = Ccl_Formata_Importacao(sCcl1, sCcl, iCclPreenchida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If Len(Trim(sCcl)) > 0 Then
                        Set objCcl = New ClassCcl
                        objCcl.sCcl = sCcl
                        lErro = CF("Ccl_Le", objCcl)
                        If lErro <> SUCESSO And lErro <> 5599 Then gError ERRO_SEM_MENSAGEM
                        If lErro <> SUCESSO Then gError 184499
                    End If
                    
                    dValor = StrParaDbl(Mid(sRegistro, 41, 15)) / 100
                    If UCase(Mid(sRegistro, 56, 1)) = "D" Then
                        dValor = -dValor
                    Else
                        dSomaCreditos = dSomaCreditos + dValor
                    End If
                    sHistorico = Trim(Mid(sRegistro, 57, 150))
                    sDocOrigem = Trim(Mid(sRegistro, 207, 30))
                                        
                    'lErro = ImportCtb_Insere_Lcto(sNomeArq, iFilialEmpresa, sCcl, bImportaLcto)
                    lErro = CF("ImportCtb_Insere_Lcto", sNomeArq, iFilialEmpresa, sCcl, bImportaLcto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                        
                    'If (Not bArqSP) Or (iFilialEmpresa = 3 And left(sCcl, 1) = "3") Or (iFilialEmpresa = 1 And left(sCcl, 1) <> "3") Then
                    If bImportaLcto Then
                        lErro = Comando_Executar(lComando, "INSERT INTO ImportCtbLctos(NumIntArq, Seq, Data, Conta, Ccl, Historico, Valor, DocOrigem) VALUES (?,?,?,?,?,?,?,?)", _
                            lNumIntArq, lQtde, dtData, sConta, sCcl, sHistorico, dValor, sDocOrigem)
                        If lErro <> AD_SQL_SUCESSO Then gError 184495
                    End If
                                        
                Case 3
                
                    If StrParaLong(Mid(sRegistro, 3, 5)) <> (lQtde + 2) Then gError 184496
'                    If Abs((StrParaDbl(Mid(sRegistro, 8, 13)) / 100) - dSomaCreditos) > DELTA_VALORMONETARIO Then gError 184497
                    
            End Select
        
        End If
                
    Loop
                
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    ImportCtb_ImportarArq1 = SUCESSO
    
    Exit Function
    
Erro_ImportCtb_ImportarArq1:

    ImportCtb_ImportarArq1 = gErr

    Select Case gErr
    
        Case 184494
             Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTB_CNPJ_FILIAL", gErr)
       
        Case 184495
             Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_IMPORTCTBLCTOS", gErr)
        
        Case 184496
             Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTCTB_QTDREG_INCOMPATIVEL", gErr, StrParaLong(Mid(sRegistro, 3, 5)), (lQtde + 2))

        Case 184497
             Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_PREENCHIDA", gErr)

        Case 184498
             Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, sConta1)

        Case 184499
             Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, sCcl1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184500)

    End Select
    
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Exit Function

End Function


