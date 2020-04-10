VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaPRJInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControlInfoAdi As Object
Dim m_objUserControl As Object
Dim m_objTela As Object

Dim sProjetoAnt As String
Dim sEtapaAnt As String

Dim sProjetoIDAnt As String
Dim sEtapaIDAnt As String

Dim glNumIntPRJ As Long
Dim glNumIntPRJEtapa As Long

Dim glNumIntPRJID As Long
Dim glNumIntPRJEtapaID As Long

Dim gcolPRJCustoReceita As Collection
Dim gcolPRJCustoReceitaInfoAdic As Collection
Dim gbConsulta As Boolean

Dim bInfoAdicIni As Boolean

Private WithEvents objEventoProjeto As AdmEvento
Attribute objEventoProjeto.VB_VarHelpID = -1
Private WithEvents objEventoProjetoInfoAdi As AdmEvento
Attribute objEventoProjetoInfoAdi.VB_VarHelpID = -1

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
    Call Inicializa_Mascara_Projeto(Projeto)
    Projeto.PromptInclude = True
    Projeto.AllowPrompt = True
End Property

Public Property Get objUserControlInfoAdi() As Object
    Set objUserControlInfoAdi = m_objUserControlInfoAdi
End Property

Public Property Set objUserControlInfoAdi(ByVal vData As Object)
    Set m_objUserControlInfoAdi = vData
    Call Inicializa_Mascara_Projeto(ProjetoInfoAdi)
    ProjetoInfoAdi.PromptInclude = True
    ProjetoInfoAdi.AllowPrompt = True
End Property

Public Property Get objTela() As Object
    Set objTela = m_objTela
End Property

Public Property Set objTela(ByVal vData As Object)
    Set m_objTela = vData
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get ProjetoInfoAdi() As Object
     Set ProjetoInfoAdi = objUserControlInfoAdi.Controls("Projeto")
End Property

Public Property Get EtapaInfoAdi() As Object
     Set EtapaInfoAdi = objUserControlInfoAdi.Controls("Etapa")
End Property

Public Property Get PRJDescInfoAdi() As Object
     Set PRJDescInfoAdi = objUserControlInfoAdi.Controls("PRJDesc")
End Property

Public Property Get PRJCliInfoAdi() As Object
     Set PRJCliInfoAdi = objUserControlInfoAdi.Controls("PRJCli")
End Property

Public Property Get PRJRespInfoAdi() As Object
     Set PRJRespInfoAdi = objUserControlInfoAdi.Controls("PRJResp")
End Property

Public Property Get PRJDtIniInfoAdi() As Object
     Set PRJDtIniInfoAdi = objUserControlInfoAdi.Controls("PRJDtIni")
End Property

Public Property Get PRJDtFimInfoAdi() As Object
     Set PRJDtFimInfoAdi = objUserControlInfoAdi.Controls("PRJDtFim")
End Property

Public Function Projeto_Validate(Cancel As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As New ClassItensPRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim bPossuiDocOriginal As Boolean
Dim objNF As New ClassNFiscal
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Projeto_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Or sEtapaAnt <> SCodigo_Extrai(Etapa.Text) Then

        If Len(Trim(Projeto.clipText)) > 0 Then
    
            'Obtém os itens da tela
            lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
            If lErro <> SUCESSO Then gError 181751
            
            Set objItemPRJCR = colItensPRJCR.Item(1)
            
            bPossuiDocOriginal = False
            If objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITREC Then
            
                lErro = Trata_Titulos(objItemPRJCR.iTipoOrigem, bPossuiDocOriginal, objNF)
                If lErro <> SUCESSO Then gError 181841
            
            End If
            
            If bPossuiDocOriginal Then gError 181842
            
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189099
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181669
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181670
            
            If sProjetoAnt <> Projeto.Text Then
                Call Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185832
            
            End If
                    
            For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
                If Len(Trim(gcolPRJCustoReceita.Item(iIndice).sItem)) = 0 Then
                    gcolPRJCustoReceita.Remove iIndice
                End If
            Next
            
            'Se já existia associação de custos e receitas
            If gcolPRJCustoReceita.Count <> 0 And sEtapaAnt <> "" Then
            
                'Avisa que vai refazer
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_PRJCR_EXISTENTE")
                If vbResult = vbNo Then gError 181750
                
            End If
                       
            'Para cada item aloca 100%
            For Each objItemPRJCR In colItensPRJCR
            
                Set objPRJCR = New ClassPRJCR
            
                objPRJCR.dPercent = 1
                objPRJCR.dQuantidade = objItemPRJCR.dQuantidadeOriginal
                objPRJCR.dValor = objItemPRJCR.dValorOriginal
                objPRJCR.iCalcAuto = MARCADO
                objPRJCR.iTipoOrigem = objItemPRJCR.iTipoOrigem
                objPRJCR.sItem = objItemPRJCR.sItem
                objPRJCR.sProjeto = objProjeto.sCodigo
                objPRJCR.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objPRJCR.lNumIntDocEtapa = objEtapa.lNumIntDoc
                objPRJCR.sObservacao = objItemPRJCR.sObservacao
                
                colPRJCR.Add objPRJCR
            
            Next
            
            'Atualiza a coleção global
            For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
                gcolPRJCustoReceita.Remove iIndice
            Next
            
            For Each objPRJCR In colPRJCR
                If Len(Trim(objPRJCR.sItem)) <> 0 Then
                    gcolPRJCustoReceita.Add objPRJCR
                End If
            Next
        
            glNumIntPRJ = objProjeto.lNumIntDoc
            glNumIntPRJEtapa = objEtapa.lNumIntDoc
            
        Else
        
            glNumIntPRJ = 0
            glNumIntPRJEtapa = 0
            
            Etapa.Clear
            
        End If
        
        sProjetoAnt = Projeto.Text
        sEtapaAnt = SCodigo_Extrai(Etapa.Text)
        
    End If
    
    Projeto_Validate = SUCESSO
    
    Exit Function

Erro_Projeto_Validate:

    Projeto_Validate = gErr

    Cancel = True

    Select Case gErr
    
        Case 181669, 181750, 181841, 185832, 189099
        
        Case 181670
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case 181842
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJ_TITULO_NF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181671)

    End Select

    Exit Function

End Function

Public Function ProjetoInfoAdi_Validate(Cancel As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As New ClassItensPRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim bPossuiDocOriginal As Boolean
Dim objNF As New ClassNFiscal
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer
Dim objCliente As ClassCliente, lCliente As Long

On Error GoTo Erro_ProjetoInfoAdi_Validate

    'Se alterou o projeto
    If sProjetoIDAnt <> ProjetoInfoAdi.Text Or sEtapaIDAnt <> SCodigo_Extrai(EtapaInfoAdi.Text) Then

        If Len(Trim(ProjetoInfoAdi.clipText)) > 0 Then
    
            'Obtém os itens da tela
            lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
            If lErro <> SUCESSO Then gError 181751
            
            Set objItemPRJCR = colItensPRJCR.Item(1)
            
            bPossuiDocOriginal = False
            If objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITREC Then
            
                lErro = Trata_Titulos(objItemPRJCR.iTipoOrigem, bPossuiDocOriginal, objNF)
                If lErro <> SUCESSO Then gError 181841
            
            End If
            
            If bPossuiDocOriginal Then gError 181842
            
            lErro = Projeto_Formata(ProjetoInfoAdi.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189099
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181669
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181670
            
            lErro = Retorno_Projeto_Tela(ProjetoInfoAdi, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 181669
            
            If sProjetoIDAnt <> ProjetoInfoAdi.Text Then
                Call Trata_Etapa(objProjeto.lNumIntDoc, EtapaInfoAdi)
            End If
            
            If Len(Trim(EtapaInfoAdi.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(EtapaInfoAdi.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185832
            
            End If
                             
            For iIndice = gcolPRJCustoReceitaInfoAdic.Count To 1 Step -1
                If Len(Trim(gcolPRJCustoReceitaInfoAdic.Item(iIndice).sItem)) = 0 Then
                    gcolPRJCustoReceitaInfoAdic.Remove iIndice
                End If
            Next
            
            'Se já existia associação de custos e receitas
            If gcolPRJCustoReceitaInfoAdic.Count <> 0 And sEtapaIDAnt <> "" Then
            
                'Avisa que vai refazer
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_PRJCR_EXISTENTE")
                If vbResult = vbNo Then gError 181750
                
            End If
                       
            'Para cada item aloca 100%
            For Each objItemPRJCR In colItensPRJCR
            
                Set objPRJCR = New ClassPRJCR
            
                objPRJCR.dPercent = 1
                objPRJCR.dQuantidade = objItemPRJCR.dQuantidadeOriginal
                objPRJCR.dValor = objItemPRJCR.dValorOriginal
                objPRJCR.iCalcAuto = MARCADO
                objPRJCR.iTipoOrigem = objItemPRJCR.iTipoOrigem
                objPRJCR.sItem = objItemPRJCR.sItem
                objPRJCR.sProjeto = objProjeto.sCodigo
                objPRJCR.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objPRJCR.lNumIntDocEtapa = objEtapa.lNumIntDoc
                objPRJCR.sObservacao = objItemPRJCR.sObservacao
                
                colPRJCR.Add objPRJCR
            
            Next
            
            'Atualiza a coleção global
            For iIndice = gcolPRJCustoReceitaInfoAdic.Count To 1 Step -1
                gcolPRJCustoReceitaInfoAdic.Remove iIndice
            Next
            
            For Each objPRJCR In colPRJCR
                If Len(Trim(objPRJCR.sItem)) <> 0 Then
                    gcolPRJCustoReceitaInfoAdic.Add objPRJCR
                End If
            Next
        
            glNumIntPRJID = objProjeto.lNumIntDoc
            glNumIntPRJEtapaID = objEtapa.lNumIntDoc
            
            PRJDescInfoAdi.Caption = objProjeto.sDescricao
            
            If objEtapa.lCliente <> 0 Then
                lCliente = objEtapa.lCliente
            Else
                lCliente = objProjeto.lCliente
            End If
            
            If lCliente <> 0 Then
            
                Set objCliente = New ClassCliente
                
                objCliente.lCodigo = lCliente
                
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 181750
                
                PRJCliInfoAdi.Caption = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sRazaoSocial
           
            Else
            
                PRJCliInfoAdi.Caption = ""
            
            End If
            
            If objEtapa.sResponsavel <> "" Then
                PRJRespInfoAdi.Caption = objEtapa.sResponsavel
            Else
                PRJRespInfoAdi.Caption = objProjeto.sResponsavel
            End If
            
            If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
                PRJDtIniInfoAdi.Caption = Format(objEtapa.dtDataInicio, "dd/mm/yyyy")
            ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
                PRJDtIniInfoAdi.Caption = Format(objProjeto.dtDataInicio, "dd/mm/yyyy")
            Else
                PRJDtIniInfoAdi.Caption = ""
            End If
            
            If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
                PRJDtFimInfoAdi.Caption = Format(objEtapa.dtDataFim, "dd/mm/yyyy")
            ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
                PRJDtFimInfoAdi.Caption = Format(objProjeto.dtDataFim, "dd/mm/yyyy")
            Else
                PRJDtFimInfoAdi.Caption = ""
            End If
            
        Else
        
            glNumIntPRJID = 0
            glNumIntPRJEtapaID = 0
            
            EtapaInfoAdi.Clear
            
            PRJDescInfoAdi.Caption = ""
            PRJCliInfoAdi.Caption = ""
            PRJRespInfoAdi.Caption = ""
            PRJDtIniInfoAdi.Caption = ""
            PRJDtFimInfoAdi.Caption = ""
            
        End If
        
        sProjetoIDAnt = ProjetoInfoAdi.Text
        sEtapaIDAnt = SCodigo_Extrai(EtapaInfoAdi.Text)
        
    End If
    
    ProjetoInfoAdi_Validate = SUCESSO
    
    Exit Function

Erro_ProjetoInfoAdi_Validate:

    ProjetoInfoAdi_Validate = gErr

    Cancel = True

    Select Case gErr
    
        Case 181669, 181750, 181841, 185832, 189099
        
        Case 181670
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case 181842
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJ_TITULO_NF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181671)

    End Select

    Exit Function

End Function

Public Function BotaoProjetos_Click(Optional ByVal bConsulta As Boolean = False) As Long

Dim lErro As Long
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As ClassItensPRJCR
Dim bPossuiDocOriginal As Boolean
Dim objNF As New ClassNFiscal
Dim bTemQtd As Boolean

On Error GoTo Erro_BotaoProjetos_Click

    gbConsulta = bConsulta

    'Obtém os itens da tela
    lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
    If lErro <> SUCESSO Then gError 181703
    
    Set objItemPRJCR = colItensPRJCR.Item(1)
    
    bPossuiDocOriginal = False
    If objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITREC Then
    
        lErro = Trata_Titulos(objItemPRJCR.iTipoOrigem, bPossuiDocOriginal, objNF)
        If lErro <> SUCESSO Then gError 181839
    
    End If
    
    If bPossuiDocOriginal Then gError 181840
    
    'Atualiza os itens de cálculo automático
    Call Atualiza_Dados(colItensPRJCR)
    
    If objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_SAQUE Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_DEPOSITO Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITREC Or objItemPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG Then
        bTemQtd = False
    Else
        bTemQtd = True
    End If

    'Abre tela de associação
    If bInfoAdicIni Then
        lErro = Chama_Tela_Modal("ProjetosCR", objUserControl, gcolPRJCustoReceitaInfoAdic, colItensPRJCR, bConsulta, bTemQtd)
    Else
        lErro = Chama_Tela_Modal("ProjetosCR", objUserControl, gcolPRJCustoReceita, colItensPRJCR, bConsulta, bTemQtd)
    End If
    If lErro <> SUCESSO Then gError 181677

    BotaoProjetos_Click = SUCESSO

    Exit Function

Erro_BotaoProjetos_Click:

    BotaoProjetos_Click = gErr

    Select Case gErr
    
        Case 181677, 181703, 181839
        
        Case 181840
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJ_TITULO_NF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Function
    
End Function

Public Function LabelProjeto_Click() As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Projeto.clipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189100

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Código")

    LabelProjeto_Click = SUCESSO

    Exit Function

Erro_LabelProjeto_Click:

    LabelProjeto_Click = gErr

    Select Case gErr
    
        Case 189100

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function
    
End Function

Private Sub objEventoProjeto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoProjeto_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189131
    
    Call objTela.Projeto_Validate(bSGECancelDummy)
    
    objUserControl.Show

    Exit Sub

Erro_objEventoProjeto_evSelecao:

    Select Case gErr
    
        Case 189131

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Sub

End Sub

Public Function LabelProjetoInfoAdi_Click() As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjetoInfoAdi_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(ProjetoInfoAdi.clipText)) <> 0 Then

        lErro = Projeto_Formata(ProjetoInfoAdi.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189100

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela_Modal("ProjetosLista", colSelecao, objProjeto, objEventoProjetoInfoAdi, , "Código")

    LabelProjetoInfoAdi_Click = SUCESSO

    Exit Function

Erro_LabelProjetoInfoAdi_Click:

    LabelProjetoInfoAdi_Click = gErr

    Select Case gErr
    
        Case 189100

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function
    
End Function

Private Sub objEventoProjetoInfoAdi_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoProjetoInfoAdi_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(ProjetoInfoAdi, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189131
    
    Call objUserControlInfoAdi.Projeto_Validate(bSGECancelDummy)
    
    objUserControlInfoAdi.Show

    Exit Sub

Erro_objEventoProjetoInfoAdi_evSelecao:

    Select Case gErr
    
        Case 189131

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Sub

End Sub

Public Function Move_Tela_Memoria(objProjetoInfo As ClassProjetoInfo, ByVal iTipoOrigem As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPRJCR As ClassPRJCR
Dim colItensPRJCR As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    Set objProjetoInfo = New ClassProjetoInfo
    
    objProjetoInfo.iTipoOrigem = iTipoOrigem
    objProjetoInfo.lNumIntDocEtapa = glNumIntPRJEtapa
    objProjetoInfo.lNumIntDocPRJ = glNumIntPRJ
    
    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189101

    objProjetoInfo.sProjeto = sProjeto
    objProjetoInfo.sEtapa = SCodigo_Extrai(Etapa.Text)

    'Obtém os itens da tela
    lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
    If lErro <> SUCESSO Then gError 181752
    
    'Atualiza os itens de cálculo automático
    lErro = Atualiza_Dados(colItensPRJCR)
    If lErro <> SUCESSO Then gError 181753
     
    For Each objPRJCR In gcolPRJCustoReceita
        objProjetoInfo.colPRJCustoReceitaReal.Add objPRJCR
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 181752, 181753, 189101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Traz_Dados_Tela(ByVal objPRJInfo As ClassProjetoInfo, ByVal lNumIntDocOrigem As Long, ByVal iTipoOrigem As Integer, Optional ByVal sCodOP As String = "", Optional ByVal iFilialEmpresa As Integer = 0) As Long

Dim lErro As Long
Dim objProjetoInfo As ClassProjetoInfo
Dim objPRJCRReal As ClassPRJCR
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Traz_Dados_Tela

    If objPRJInfo Is Nothing Then
    
        lErro = Obtem_ProjetoInfo(objProjetoInfo, lNumIntDocOrigem, iTipoOrigem, sCodOP, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 181674
        
    Else
    
        Set objProjetoInfo = objPRJInfo
    
    End If
    
    objProjeto.lNumIntDoc = objProjetoInfo.lNumIntDocPRJ
    
    If objProjeto.lNumIntDoc <> 0 Then
    
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181679
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181780
        
    End If
    
    glNumIntPRJ = objProjeto.lNumIntDoc
    
    objEtapa.lNumIntDoc = objProjetoInfo.lNumIntDocEtapa
    
    If objEtapa.lNumIntDoc <> 0 Then
    
        lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185833
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185834
        
    End If
    
    glNumIntPRJEtapa = objEtapa.lNumIntDoc
    
    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189132
    
    sProjetoAnt = Projeto.Text
    
    sEtapaAnt = objEtapa.sCodigo
    
    Call Trata_Etapa(glNumIntPRJ, Etapa)
    
    Call CF("SCombo_Seleciona2", Etapa, sEtapaAnt)
    
    For Each objPRJCRReal In objProjetoInfo.colPRJCustoReceitaReal
        gcolPRJCustoReceita.Add objPRJCRReal
    Next
    
    Traz_Dados_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Tela:

    Traz_Dados_Tela = gErr

    Select Case gErr
    
        Case 181674, 181679, 185833, 189132
        
        Case 181780
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
        
        Case 185834
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Obtem_ProjetoInfo(objProjetoInfo As ClassProjetoInfo, ByVal lNumIntDocOrigem As Long, ByVal iTipoOrigem As Integer, Optional ByVal sCodOP As String = "", Optional ByVal iFilialEmpresa As Integer = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Obtem_ProjetoInfo

    Set objProjetoInfo = New ClassProjetoInfo

    objProjetoInfo.iTipoOrigem = iTipoOrigem
    objProjetoInfo.lNumIntDocOrigem = lNumIntDocOrigem
    objProjetoInfo.sCodigoOP = sCodOP
    objProjetoInfo.iFilialEmpresa = iFilialEmpresa

    'Le as associação gravadas no BD para esse tipo de documento
    lErro = CF("ProjetoInfo_Le", objProjetoInfo)
    If lErro <> SUCESSO Then gError 181838
            
    Obtem_ProjetoInfo = SUCESSO

    Exit Function

Erro_Obtem_ProjetoInfo:

    Obtem_ProjetoInfo = gErr

    Select Case gErr
    
        Case 181838

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181839)

    End Select

    Exit Function

End Function

Public Function Atualiza_Dados(ByVal colItensPRJCR As Collection) As Long

Dim lErro As Long
Dim objPRJCR As ClassPRJCR
Dim objItemPRJCR As ClassItensPRJCR
Dim iIndice As Integer
Dim colPRJCR As Collection

On Error GoTo Erro_Atualiza_Dados

    If bInfoAdicIni Then
        Set colPRJCR = gcolPRJCustoReceitaInfoAdic
    Else
        Set colPRJCR = gcolPRJCustoReceita
    End If

    For iIndice = colPRJCR.Count To 1 Step -1
        If Len(Trim(colPRJCR.Item(iIndice).sItem)) = 0 Then
            colPRJCR.Remove iIndice
        End If
    Next

    'Para cada associação
    For Each objPRJCR In colPRJCR

        'Verifica se o cálculo é automático
        If objPRJCR.iCalcAuto = MARCADO Then
        
            'Encontra o item da tela
            For Each objItemPRJCR In colItensPRJCR
                If objItemPRJCR.sItem = objPRJCR.sItem Then
                    Exit For
                End If
            Next
            
            'Recalcula quantidade e valor
            objPRJCR.dQuantidade = objItemPRJCR.dQuantidadeOriginal * objPRJCR.dPercent
            objPRJCR.dValor = objItemPRJCR.dValorOriginal * objPRJCR.dPercent
            objPRJCR.iTipoOrigem = objItemPRJCR.iTipoOrigem
            objPRJCR.iTipoValor = objItemPRJCR.iNegativo
            
            If Len(Trim(objPRJCR.sObservacao)) = 0 Then
                objPRJCR.sObservacao = objItemPRJCR.sObservacao
            End If
    
        End If
    
    Next
    
    Atualiza_Dados = SUCESSO

    Exit Function

Erro_Atualiza_Dados:

    Atualiza_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Valida_Dados() As Long

Dim lErro As Long
Dim objItemPRJCR As ClassItensPRJCR
Dim dValor As Double
Dim dQuantidade As Double
Dim objPRJCR As ClassPRJCR
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim sProjeto As String, iProjetoPreenchido As Integer

On Error GoTo Erro_Valida_Dados

    If gobjFAT.iPRJObrigatorio = MARCADO Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 181752
        
        If iProjetoPreenchido = DESMARCADO Then gError 213116
    
    End If

    'Obtém os itens da tela
    lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
    If lErro <> SUCESSO Then gError 181752

    For Each objItemPRJCR In colItensPRJCR
    
        dValor = 0
        dQuantidade = 0
        
        For Each objPRJCR In gcolPRJCustoReceita
    
            If objItemPRJCR.sItem = objPRJCR.sItem Then
                dValor = dValor + objPRJCR.dValor
                dQuantidade = dQuantidade + objPRJCR.dQuantidade
            End If
    
        Next
        
        If objItemPRJCR.dValorOriginal - dValor < -DELTA_VALORMONETARIO Then gError 181757
        If objItemPRJCR.dQuantidadeOriginal - dQuantidade < -DELTA_VALORMONETARIO Then gError 181758
        
        If gobjFAT.iTipoValidacaoPRJ = PRJ_TIPO_VALID_VLR_MENOR_GRAVACAO Or gobjFAT.iTipoValidacaoPRJ = PRJ_TIPO_VALID_VLR_MENOR_AMBOS Then
        
            If dValor <> 0 Or dQuantidade <> 0 Then
                If Abs(objItemPRJCR.dValorOriginal - dValor) > DELTA_VALORMONETARIO Then
                
                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_VALOR_DIFERE_ITEMPRJ", objItemPRJCR.sItem, Format(objItemPRJCR.dValorOriginal, "STANDARD"), Format(dValor, "STANDARD"))
                    If vbResult = vbNo Then gError 181759
                
                End If
            
                If Abs(objItemPRJCR.dQuantidadeOriginal - dQuantidade) > DELTA_VALORMONETARIO Then
                
                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_QTD_DIFERE_ITEMPRJ", objItemPRJCR.sItem, Formata_Estoque(objItemPRJCR.dQuantidadeOriginal), Formata_Estoque(dQuantidade))
                    If vbResult = vbNo Then gError 181760
                
                End If
                
            End If
            
        End If
        
    Next
    
    Valida_Dados = SUCESSO

    Exit Function

Erro_Valida_Dados:

    Valida_Dados = gErr

    Select Case gErr
    
        Case 181757
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PROJETO_MAIOR_VALOR_ITEM", gErr, objItemPRJCR.sItem, Format(objItemPRJCR.dValorOriginal, "STANDARD"), Format(dValor, "STANDARD"))
        
        Case 181758
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_PROJETO_MAIOR_QTD_ITEM", gErr, objItemPRJCR.sItem, Formata_Estoque(objItemPRJCR.dQuantidadeOriginal), Formata_Estoque(dQuantidade))
        
        Case 181759, 181760, 181752
        
        Case 213116
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Remove_Linha(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_Remove_Linha

    For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
    
        Set objPRJCR = gcolPRJCustoReceita.Item(iIndice)

        If objPRJCR.sItem = CStr(iLinha) Then
            gcolPRJCustoReceita.Remove iIndice
        End If

    Next
    
    For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
    
        Set objPRJCR = gcolPRJCustoReceita.Item(iIndice)
        
        If IsNumeric(objPRJCR.sItem) Then
            If StrParaInt(objPRJCR.sItem) > iLinha Then
                objPRJCR.sItem = CStr(StrParaInt(objPRJCR.sItem) - 1)
            End If
        End If

    Next
    
    Remove_Linha = SUCESSO

    Exit Function

Erro_Remove_Linha:

    Remove_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Remove_Linha2(ByVal objPRJInfo As ClassProjetoInfo, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_Remove_Linha2

    For iIndice = objPRJInfo.colPRJCustoReceitaReal.Count To 1 Step -1
    
        Set objPRJCR = objPRJInfo.colPRJCustoReceitaReal.Item(iIndice)

        If objPRJCR.sItem = CStr(iLinha) Then
            objPRJInfo.colPRJCustoReceitaReal.Remove iIndice
        End If

    Next
    
    For iIndice = objPRJInfo.colPRJCustoReceitaReal.Count To 1 Step -1
    
        Set objPRJCR = objPRJInfo.colPRJCustoReceitaReal.Item(iIndice)
        
        If IsNumeric(objPRJCR.sItem) Then
            If StrParaInt(objPRJCR.sItem) > iLinha Then
                objPRJCR.sItem = CStr(StrParaInt(objPRJCR.sItem) - 1)
            End If
        End If

    Next
    
    Remove_Linha2 = SUCESSO

    Exit Function

Erro_Remove_Linha2:

    Remove_Linha2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Troca_Linha(ByVal iLinhaAtual As Integer, ByVal iLinhaNova As Integer) As Long

Dim lErro As Long, iIndice As Integer
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_Troca_Linha

    lErro = Colecao_Altera_Ordem_Itens(gcolPRJCustoReceita, iLinhaAtual, iLinhaNova)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iIndice = 0
    For Each objPRJCR In gcolPRJCustoReceita
        If IsNumeric(objPRJCR.sItem) Then
            iIndice = iIndice + 1
            objPRJCR.sItem = CStr(iIndice)
        End If
    Next

    Troca_Linha = SUCESSO

    Exit Function

Erro_Troca_Linha:

    Troca_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Adiciona_Linha(ByVal iItem As Integer) As Long

Dim lErro As Long
Dim objPRJCR As ClassPRJCR
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Adiciona_Linha

    If Len(Trim(Projeto.clipText)) > 0 Then

        Set objPRJCR = New ClassPRJCR
    
        objPRJCR.dPercent = 1
        objPRJCR.iCalcAuto = MARCADO
        objPRJCR.sItem = CStr(iItem)
        
        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189102

        objPRJCR.sProjeto = sProjeto
        objPRJCR.lNumIntDocPRJ = glNumIntPRJ
        objPRJCR.lNumIntDocEtapa = glNumIntPRJEtapa
        
        gcolPRJCustoReceita.Add objPRJCR

    End If
    
    Adiciona_Linha = SUCESSO

    Exit Function

Erro_Adiciona_Linha:

    Adiciona_Linha = gErr

    Select Case gErr
    
        Case 189102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Function

End Function

Public Function Trata_Titulos(ByVal iTipoOrigem As Integer, bPossuiDocOriginal As Boolean, ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Titulos

    Select Case iTipoOrigem
    
        Case PRJ_CR_TIPO_TITPAG
            
            lErro = Trata_TitulosPag(bPossuiDocOriginal, objNF)
            If lErro <> SUCESSO Then gError 181844
        
        Case PRJ_CR_TIPO_TITREC
        
            lErro = Trata_TitulosRec(bPossuiDocOriginal, objNF)
            If lErro <> SUCESSO Then gError 181845
        
    End Select
    
    Trata_Titulos = SUCESSO

    Exit Function

Erro_Trata_Titulos:

    Trata_Titulos = gErr

    Select Case gErr
    
        Case 181844, 181845

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181846)

    End Select

    Exit Function

End Function

Public Function Trata_TitulosPag(bPossuiDocOriginal As Boolean, ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim objTituloPag As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim sTela As String

On Error GoTo Erro_Trata_TitulosPag

    objFornecedor.sNomeReduzido = objUserControl.Controls("Fornecedor").Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 181849
    
    objTituloPag.lFornecedor = objFornecedor.lCodigo
    objTituloPag.iFilial = Codigo_Extrai(objUserControl.Controls("Filial").Text)
    objTituloPag.sSiglaDocumento = SCodigo_Extrai(objUserControl.Controls("Tipo").Text)
    If Not gbConsulta Then
        objTituloPag.lNumTitulo = StrParaLong(objUserControl.Controls("NumTitulo").Text)
    Else
        objTituloPag.lNumTitulo = StrParaLong(objUserControl.Controls("NumeroTitulo").Text)
    End If
    objTituloPag.iFilialEmpresa = giFilialEmpresa
    objTituloPag.dtDataEmissao = StrParaDate(objUserControl.Controls("DataEmissao").Text)
    
'    'Procura o Titulo (Baixados ou não)
'    lErro = CF("TituloPag_Le_Todos", objTituloPag)
'    If lErro <> SUCESSO Then gError 181850

    'Lê em Títulos a Pagar
    lErro = CF("TituloPagar_Le_Numero", objTituloPag)
    If lErro <> SUCESSO And lErro <> 18551 Then gError 181850
    
    'Se não encontrou
    If lErro = 18551 Then
    
        'Procura em Títulos a Receber Baixados
        lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPag)
        If lErro <> SUCESSO And lErro <> 18556 Then gError 181850

    End If
                
    If lErro = SUCESSO Then
                
        'Le o Documento que originou este Título
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPag.lNumIntDoc, CPR_TITULO_PAGAR, objNF, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 181851
        
        If lErro = SUCESSO Then bPossuiDocOriginal = True
        
    End If
    
    Trata_TitulosPag = SUCESSO

    Exit Function

Erro_Trata_TitulosPag:

    Trata_TitulosPag = gErr

    Select Case gErr
    
        Case 181849 To 181851

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181847)

    End Select

    Exit Function

End Function

Public Function Trata_TitulosRec(bPossuiDocOriginal As Boolean, ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim objTituloRec As New ClassTituloReceber
Dim sTela As String
Dim objCliente As New ClassCliente

On Error GoTo Erro_Trata_TitulosRec

    objCliente.sNomeReduzido = objUserControl.Controls("Cliente").Text

    'Lê o codigo do Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 181852

    objTituloRec.lCliente = objCliente.lCodigo
    objTituloRec.iFilial = Codigo_Extrai(objUserControl.Controls("Filial").Text)
    objTituloRec.sSiglaDocumento = SCodigo_Extrai(objUserControl.Controls("Tipo").Text)
    If Not gbConsulta Then
        objTituloRec.lNumTitulo = StrParaLong(objUserControl.Controls("NumTitulo").Text)
    Else
        objTituloRec.lNumTitulo = StrParaLong(objUserControl.Controls("NumeroTitulo").Text)
    End If
    objTituloRec.iFilialEmpresa = giFilialEmpresa

    'Procura o Titulo (Baixados ou não)
    lErro = CF("TitulosRec_Le_Todos", objTituloRec)
    If lErro <> SUCESSO And lErro <> 58997 Then gError 181853

    'Le o documento que originou este Título
    lErro = CF("Titulo_Le_DocumentoOriginal", objTituloRec.lNumIntDoc, CPR_TITULO_RECEBER, objNF, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError 181854

    If lErro = SUCESSO Then bPossuiDocOriginal = True
   
    Trata_TitulosRec = SUCESSO

    Exit Function

Erro_Trata_TitulosRec:

    Trata_TitulosRec = gErr

    Select Case gErr
    
        Case 181852 To 181854

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181848)

    End Select

    Exit Function

End Function

Public Function Trata_Etapa(ByVal lNumIntDocPRJ As Long, ByVal objCombo As Object) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos

On Error GoTo Erro_Trata_Etapa
    
    If lNumIntDocPRJ <> 0 Then

        objProjeto.lNumIntDoc = lNumIntDocPRJ
        
        objCombo.AddItem ""
    
        lErro = CF("CarregaCombo_Etapas", objProjeto, objCombo)
        If lErro <> SUCESSO Then gError 185234
    
    Else
    
        objCombo.Clear
        
    End If

    Trata_Etapa = SUCESSO

    Exit Function

Erro_Trata_Etapa:

    Trata_Etapa = gErr

    Select Case gErr
    
        Case 185234

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185235)

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    Set objEventoProjeto = New AdmEvento
    Set objEventoProjetoInfoAdi = New AdmEvento
    Set gcolPRJCustoReceita = New Collection
    Set gcolPRJCustoReceitaInfoAdic = New Collection
    bInfoAdicIni = False
End Sub

Private Sub Class_Terminate()
    Etapa.Clear
    Set objEventoProjeto = Nothing
    Set objEventoProjetoInfoAdi = Nothing
    Set gcolPRJCustoReceita = Nothing
    Set gcolPRJCustoReceitaInfoAdic = Nothing
    Set m_objUserControl = Nothing
    Set m_objUserControlInfoAdi = Nothing
    Set m_objTela = Nothing
End Sub

Public Function Inicializa_InfoAdi() As Long

Dim lErro As Long, iIndice As Integer
Dim objPRJCR As ClassPRJCR, lCliente As Long
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas
Dim objCliente As New ClassCliente

On Error GoTo Erro_Inicializa_InfoAdi

    bInfoAdicIni = True
    
    sProjetoIDAnt = sProjetoAnt
    sEtapaIDAnt = sEtapaAnt
        
    glNumIntPRJID = glNumIntPRJ
    glNumIntPRJEtapaID = glNumIntPRJEtapa

    ProjetoInfoAdi.PromptInclude = False
    ProjetoInfoAdi.Text = Projeto.Text
    ProjetoInfoAdi.PromptInclude = True
    
    Call Trata_Etapa(glNumIntPRJ, EtapaInfoAdi)
    
    Call CF("SCombo_Seleciona2", EtapaInfoAdi, SCodigo_Extrai(Etapa.Text))
    
    For iIndice = gcolPRJCustoReceitaInfoAdic.Count To 1 Step -1
        gcolPRJCustoReceitaInfoAdic.Remove iIndice
    Next
    
    For Each objPRJCR In gcolPRJCustoReceita
        gcolPRJCustoReceitaInfoAdic.Add objPRJCR
    Next
    
    If glNumIntPRJID > 0 Then
    
        objProjeto.lNumIntDoc = glNumIntPRJID
        
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(EtapaInfoAdi.Text)) > 0 Then
        
            objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
            objEtapa.sCodigo = SCodigo_Extrai(EtapaInfoAdi.Text)
        
            lErro = CF("PrjEtapas_Le", objEtapa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        PRJDescInfoAdi.Caption = objProjeto.sDescricao
         
         If objEtapa.lCliente <> 0 Then
             lCliente = objEtapa.lCliente
         Else
             lCliente = objProjeto.lCliente
         End If
         
         If lCliente <> 0 Then
         
             Set objCliente = New ClassCliente
             
             objCliente.lCodigo = lCliente
             
             lErro = CF("Cliente_Le", objCliente)
             If lErro <> SUCESSO And lErro <> 12348 Then gError 181750
             
             PRJCliInfoAdi.Caption = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sRazaoSocial
        
         Else
         
             PRJCliInfoAdi.Caption = ""
         
         End If
         
         If objEtapa.sResponsavel <> "" Then
             PRJRespInfoAdi.Caption = objEtapa.sResponsavel
         Else
             PRJRespInfoAdi.Caption = objProjeto.sResponsavel
         End If
         
         If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
             PRJDtIniInfoAdi.Caption = Format(objEtapa.dtDataInicio, "dd/mm/yyyy")
         ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
             PRJDtIniInfoAdi.Caption = Format(objProjeto.dtDataInicio, "dd/mm/yyyy")
         Else
             PRJDtIniInfoAdi.Caption = ""
         End If
         
         If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
             PRJDtFimInfoAdi.Caption = Format(objEtapa.dtDataFim, "dd/mm/yyyy")
         ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
             PRJDtFimInfoAdi.Caption = Format(objProjeto.dtDataFim, "dd/mm/yyyy")
         Else
             PRJDtFimInfoAdi.Caption = ""
         End If
    End If

    Inicializa_InfoAdi = SUCESSO

    Exit Function

Erro_Inicializa_InfoAdi:

    Inicializa_InfoAdi = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185235)

    End Select

    Exit Function

End Function

Public Function Finaliza_InfoAdi() As Long
    bInfoAdicIni = False
    sProjetoIDAnt = ""
    sEtapaIDAnt = ""
    glNumIntPRJID = 0
    glNumIntPRJEtapaID = 0
End Function

Public Function Transfere_InfoAdi() As Long

Dim lErro As Long, iIndice As Integer
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_Transfere_InfoAdi

    'Atualiza a coleção global
    For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
        gcolPRJCustoReceita.Remove iIndice
    Next
    
    For Each objPRJCR In gcolPRJCustoReceitaInfoAdic
        gcolPRJCustoReceita.Add objPRJCR
    Next
       
    sProjetoAnt = sProjetoIDAnt
    sEtapaAnt = sEtapaIDAnt
        
    glNumIntPRJ = glNumIntPRJID
    glNumIntPRJEtapa = glNumIntPRJEtapaID
    
    Projeto.PromptInclude = False
    Projeto.Text = ProjetoInfoAdi.Text
    Projeto.PromptInclude = True
    
    Call Trata_Etapa(glNumIntPRJ, Etapa)
    
    Call CF("SCombo_Seleciona2", Etapa, SCodigo_Extrai(EtapaInfoAdi.Text))

    Transfere_InfoAdi = SUCESSO

    Exit Function

Erro_Transfere_InfoAdi:

    Transfere_InfoAdi = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185235)

    End Select

    Exit Function

End Function

Public Function Recalcula_Dados_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As New ClassItensPRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim bPossuiDocOriginal As Boolean
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Recalcula_Dados_Tela

    If Len(Trim(Projeto.clipText)) > 0 Then

        'Obtém os itens da tela
        lErro = objTela.Obter_ItensPRJCR(colItensPRJCR)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                   
        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrou => Erro
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181670
                    
        If Len(Trim(Etapa.Text)) > 0 Then
        
            objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
            objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
        
            lErro = CF("PrjEtapas_Le", objEtapa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        End If
                    
        For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
            If Len(Trim(gcolPRJCustoReceita.Item(iIndice).sItem)) = 0 Then
                gcolPRJCustoReceita.Remove iIndice
            End If
        Next
                   
        'Para cada item aloca 100%
        For Each objItemPRJCR In colItensPRJCR
        
            Set objPRJCR = New ClassPRJCR
        
            objPRJCR.dPercent = 1
            objPRJCR.dQuantidade = objItemPRJCR.dQuantidadeOriginal
            objPRJCR.dValor = objItemPRJCR.dValorOriginal
            objPRJCR.iCalcAuto = MARCADO
            objPRJCR.iTipoOrigem = objItemPRJCR.iTipoOrigem
            objPRJCR.sItem = objItemPRJCR.sItem
            objPRJCR.sProjeto = objProjeto.sCodigo
            objPRJCR.lNumIntDocPRJ = objProjeto.lNumIntDoc
            objPRJCR.lNumIntDocEtapa = objEtapa.lNumIntDoc
            objPRJCR.sObservacao = objItemPRJCR.sObservacao
            
            colPRJCR.Add objPRJCR
        
        Next
        
        'Atualiza a coleção global
        For iIndice = gcolPRJCustoReceita.Count To 1 Step -1
            gcolPRJCustoReceita.Remove iIndice
        Next
        
        For Each objPRJCR In colPRJCR
            If Len(Trim(objPRJCR.sItem)) <> 0 Then
                gcolPRJCustoReceita.Add objPRJCR
            End If
        Next
            
    End If
    
    Recalcula_Dados_Tela = SUCESSO
    
    Exit Function

Erro_Recalcula_Dados_Tela:

    Recalcula_Dados_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 181670
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181671)

    End Select

    Exit Function

End Function
