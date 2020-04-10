VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RelTelaAux_AbrirInt Lib "ADCUSR.DLL" Alias "RelTelaAux_Abrir" (ByVal sNomeArq As String) As Long
Private Declare Function RelTelaAux_FecharInt Lib "ADCUSR.DLL" Alias "RelTelaAux_Fechar" (ByVal lRelTelaAux As Long) As Long

Private Declare Function RelTelaAux_IncCpoTextoInt Lib "ADCUSR.DLL" Alias "RelTelaAux_IncCpoTexto" (ByVal lRelTelaAux As Long, ByVal sTexto As String, ByVal iTam As Integer) As Long
Private Declare Function RelTelaAux_IncCpoNumInt Lib "ADCUSR.DLL" Alias "RelTelaAux_IncCpoNum" (ByVal lRelTelaAux As Long, ByVal dblNum As Double) As Long
Private Declare Function RelTelaAux_IncCpoDataInt Lib "ADCUSR.DLL" Alias "RelTelaAux_IncCpoData" (ByVal lRelTelaAux As Long, ByVal dbl_data_jul As Double) As Long

'Para apoio a execucao de relatorios tipo "printscreen"
'Este relatorios sao executados a partir de um arquivo temporario criado a partir dos registros passados p/esta classe
'A estrutura do registro associado a um relatorio está decrita na tabela RelTelaCampos no dicdados.

Private mvarsCodRel As String
Private sNomeArqDados As String
Private mvarcolCampos As ColRelTelaCampos
Private mvarlRelTelaAux As Long 'ponteiro p/acesso a funcoes em C, de apoio

Private objRelOpcoes As New AdmRelOpcoes

Private Const RELTELA_CAMPO_TIPO_TEXTO = 0
Private Const RELTELA_CAMPO_TIPO_NUM = 1
Private Const RELTELA_CAMPO_TIPO_DATA = 2

Private Sub Class_Initialize()
    
    Set mvarcolCampos = New ColRelTelaCampos

End Sub

Private Sub Class_Terminate()
    
    Set mvarcolCampos = Nothing
    
    If mvarlRelTelaAux <> 0 Then Call RelTelaAux_Fechar(mvarlRelTelaAux)
    
End Sub

Public Function Iniciar(sCodRel As String) As Long
'Carrega a definicao dos campos p/uma colecao e prepara para receber os registros

Dim lErro As Long

On Error GoTo Erro_Iniciar

    mvarsCodRel = sCodRel
    
    'Carrega a definicao dos campos
    lErro = CF("RelTela_LeCampos", sCodRel, mvarcolCampos)
    If lErro <> SUCESSO Then Error 41564
    
    'inicializar estrutura de apoio
    lErro = RelTelaAux_Abrir(mvarlRelTelaAux, sNomeArqDados)
    If lErro <> SUCESSO Then Error 41571
    
    Iniciar = SUCESSO
     
    Exit Function
    
Erro_Iniciar:

    Iniciar = Err
     
    Select Case Err
          
        Case 41564, 41571
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153350)
     
    End Select
     
    Exit Function

End Function

Public Function IncluirRegistro(colValoresDosCampos As Collection) As Long
'inclui um registro contendo para cada campo um elemento da colecao
'verifica o tipo dos valores passados com a estrutura lida no dicdados
Dim lErro As Long, vVal As Variant, iIndice As Integer, objRelTelaCampo As ClassRelTelaCampo, sTemp As String
On Error GoTo Erro_IncluirRegistro

    If colValoresDosCampos.Count <> mvarcolCampos.Count Then Error 41572
    
    For iIndice = 1 To colValoresDosCampos.Count
    
        vVal = colValoresDosCampos.Item(iIndice)
        
        Set objRelTelaCampo = mvarcolCampos.Item(iIndice)
    
        'validar tipo do dado passado com o informado no dicdados
        Select Case objRelTelaCampo.iTipo
        
            Case RELTELA_CAMPO_TIPO_DATA
                lErro = RelTelaAux_IncCpoData(mvarlRelTelaAux, CDate(vVal))
                If lErro <> SUCESSO Then Error 41573
            
            Case RELTELA_CAMPO_TIPO_TEXTO
                sTemp = CStr(vVal)
                If Len(sTemp) > objRelTelaCampo.iTamanho Then sTemp = Left(sTemp, objRelTelaCampo.iTamanho)
                lErro = RelTelaAux_IncCpoTexto(mvarlRelTelaAux, sTemp, objRelTelaCampo.iTamanho)
                If lErro <> SUCESSO Then Error 41575
            
            Case RELTELA_CAMPO_TIPO_NUM
                lErro = RelTelaAux_IncCpoNum(mvarlRelTelaAux, CDbl(vVal))
                If lErro <> SUCESSO Then Error 41576
            
            Case Else
                Error 41577
                
        End Select
        
    Next
    
    IncluirRegistro = SUCESSO
     
    Exit Function
    
Erro_IncluirRegistro:

    IncluirRegistro = Err
     
    Select Case Err
        
        Case 41572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELTELA_NUM_CAMPOS", Err)
        
        Case 41573, 41575, 41576
        
        Case 41574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELTELA_TAM_CAMPO", Err)
        
        Case 41577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELTELA_TIPO_CAMPO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153351)
     
    End Select
     
    Exit Function

End Function

Public Function ExecutarRel(Optional sNomeTsk As String, Optional sParam1 As String, Optional sValParam1 As String, Optional sParam2 As String, Optional sValParam2 As String, Optional sParam3 As String, Optional sValParam3 As String, Optional sParam4 As String, Optional sValParam4 As String, Optional sParam5 As String, Optional sValParam5 As String, Optional sParam6 As String, Optional sValParam6 As String) As Long
'executa o relatorio
        
Dim lErro As Long, objRelatorio As New AdmRelatorio
On Error GoTo Erro_ExecutarRel

    lErro = RelTelaAux_Fechar(mvarlRelTelaAux)
    If lErro <> SUCESSO Then Error 41587
    
    If sParam1 <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro(sParam1, sValParam1)
        If lErro <> AD_BOOL_TRUE Then Error 59581

        If sParam2 <> "" Then
            
            lErro = objRelOpcoes.IncluirParametro(sParam2, sValParam2)
            If lErro <> AD_BOOL_TRUE Then Error 59582
            
            If sParam3 <> "" Then
                
                lErro = objRelOpcoes.IncluirParametro(sParam3, sValParam3)
                If lErro <> AD_BOOL_TRUE Then Error 59583
                
                If sParam4 <> "" Then
                    
                    lErro = objRelOpcoes.IncluirParametro(sParam4, sValParam4)
                    If lErro <> AD_BOOL_TRUE Then Error 59584
                    
                    If sParam5 <> "" Then
                        
                        lErro = objRelOpcoes.IncluirParametro(sParam5, sValParam5)
                        If lErro <> AD_BOOL_TRUE Then Error 59585
                        
                    End If
                
                End If
            
            End If
        
        End If
    
    End If
    
    'manda executar passando nome do arquivo como @
    lErro = objRelatorio.ExecutarDireto(mvarsCodRel, "", 0, sNomeTsk, "AARQREL", sNomeArqDados, sParam1, sValParam1, sParam2, sValParam2, sParam3, sValParam3, sParam4, sValParam4, sParam5, sValParam5)
    If lErro <> SUCESSO Then Error 41588

    ExecutarRel = SUCESSO
     
    Exit Function
    
Erro_ExecutarRel:

    ExecutarRel = Err
     
    Select Case Err
          
        Case 41587, 41588, 59581, 59582, 59583, 59584, 59585, 59586
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153352)
     
    End Select
     
    Exit Function

End Function

'*****************************************************************
'*****************************************************************
'funcoes para interface com dll em VC
'*****************************************************************
'*****************************************************************
Function RelTelaAux_Abrir(lRelTelaAux As Long, sNomeArq As String) As Long
Dim lErro As Long
On Error GoTo Erro_RelTelaAux_Abrir

    sNomeArq = String(255, 0)
    
    lRelTelaAux = RelTelaAux_AbrirInt(sNomeArq)
    If lRelTelaAux = 0 Then Error 41568
    
    sNomeArq = StringZ(sNomeArq)
    
    RelTelaAux_Abrir = SUCESSO
     
    Exit Function
    
Erro_RelTelaAux_Abrir:

    RelTelaAux_Abrir = Err
     
    Select Case Err
          
        Case 41568
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153353)
     
    End Select
     
    Exit Function

End Function

Function RelTelaAux_Fechar(lRelTelaAux As Long) As Long
'fecha arquivo e libera recursos alocados
Dim lErro As Long
On Error GoTo Erro_RelTelaAux_Fechar

    lErro = RelTelaAux_FecharInt(lRelTelaAux)
    If lErro <> AD_BOOL_TRUE Then Error 41570
    
    lRelTelaAux = 0
    
    RelTelaAux_Fechar = SUCESSO
     
    Exit Function
    
Erro_RelTelaAux_Fechar:

    RelTelaAux_Fechar = Err
     
    Select Case Err
          
        Case 41570
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153354)
     
    End Select
     
    Exit Function
    
End Function

Function RelTelaAux_IncCpoTexto(lRelTelaAux As Long, sTexto As String, iTamMax As Integer) As Long
Dim lErro As Long
On Error GoTo Erro_RelTelaAux_IncCpoTexto

    lErro = RelTelaAux_IncCpoTextoInt(lRelTelaAux, sTexto, iTamMax)
    If lErro <> AD_BOOL_TRUE Then Error 41578
    
    RelTelaAux_IncCpoTexto = SUCESSO
     
    Exit Function
    
Erro_RelTelaAux_IncCpoTexto:

    RelTelaAux_IncCpoTexto = Err
     
    Select Case Err
          
        Case 41578
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153355)
     
    End Select
     
    Exit Function

End Function

Function RelTelaAux_IncCpoNum(lRelTelaAux As Long, dNum As Double) As Long
Dim lErro As Long
On Error GoTo Erro_RelTelaAux_IncCpoNum

    lErro = RelTelaAux_IncCpoNumInt(lRelTelaAux, dNum)
    If lErro <> AD_BOOL_TRUE Then Error 41579
    
    RelTelaAux_IncCpoNum = SUCESSO
     
    Exit Function
    
Erro_RelTelaAux_IncCpoNum:

    RelTelaAux_IncCpoNum = Err
     
    Select Case Err
          
        Case 41579
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153356)
     
    End Select
     
    Exit Function
    
End Function

Function RelTelaAux_IncCpoData(lRelTelaAux As Long, dtData As Date) As Long
Dim lErro As Long
On Error GoTo Erro_RelTelaAux_IncCpoData

    lErro = RelTelaAux_IncCpoDataInt(lRelTelaAux, CDbl(dtData))
    If lErro <> AD_BOOL_TRUE Then Error 41580
    
    RelTelaAux_IncCpoData = SUCESSO
     
    Exit Function
    
Erro_RelTelaAux_IncCpoData:

    RelTelaAux_IncCpoData = Err
     
    Select Case Err
          
        Case 41581
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153357)
     
    End Select
     
    Exit Function

End Function
