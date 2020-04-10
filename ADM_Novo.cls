VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const JONES_280115 = 0

'Alteração Daniel em 07/11/02
Const MAX_NOME_CLASSE_CTL = 256
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Fim

Private Declare Function Rotina_Erro_Carregar Lib "ADCRTL.DLL" (ByVal sTipoErro As String, ByVal lLocalErro As Long, ByVal lpMsgErro As String, ByVal sParam1 As String, ByVal sParam2 As String, ByVal sParam3 As String, ByVal sParam4 As String, ByVal sParam5 As String, ByVal sParam6 As String, ByVal sParam7 As String, ByVal sParam8 As String, ByVal sParam9 As String, ByVal sParam10 As String) As Long

'para sistema e outras classes de adcusr.dll
Private Declare Function Sistema_CriarInt Lib "ADCUSR.DLL" Alias "AD_Sistema_Criar" () As Long
Private Declare Function Sistema_QueryUnloadInt Lib "ADCUSR.DLL" Alias "AD_Sistema_QueryUnload" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sistema_Criar_Batch_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_Criar_Batch" (ByVal sNomeArqParam As String, lCodEmp As Long, ByVal sNomeEmp As String, iCodFilial As Integer, ByVal sNomeFilial As String, dtDataAtual As Date) As Long
Private Declare Function Sistema_ExecutarBatch_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_ExecutarBatch" (ByVal lID_Sistema As Long, ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal sNomeFilial As String, ByVal dtDataAtual As Date) As Long
Private Declare Function Sistema_RegistrarOcorrencia Lib "ADCUSR.DLL" Alias "AD_Sistema_RegistrarOcorrencia" (ByVal lID_Sistema As Long, ByVal sTexto As String) As Long
Private Declare Function Reset_Contab_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_ResetContab" (ByVal lID_Sistema As Long) As Long
Private Declare Function Reset_Fest_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_ResetFest" (ByVal lID_Sistema As Long) As Long

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

' Inicializa adcusr.dll
'Retorno:
'    ponteiro p/sistema (long no VB) ou NULL p/erro.

Private Declare Function Sistema_DestruirInt Lib "ADCUSR.DLL" Alias "AD_Sistema_Destruir" (ByVal lID_Sistema As Long) As Long
' Finaliza adcusr.dll
'Retorno:
'    AD_BOOL_TRUE p/sucesso

Private Declare Function Sist_ObterAcesso Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterAcesso" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterConDic Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterConDic" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterConDicBrowse Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterConDicBrowse" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterConRel Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterConRel" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterConBrowse Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterConBrowse" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterMascara Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterMascara" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterMascaraProd Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterMascaraProd" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterConexao Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterConexao" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterDicDados Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterDicDados" (ByVal lID_Sistema As Long) As Long
Private Declare Function Sist_ObterTipoVersao Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterTipoVersao" (ByVal lID_Sistema As Long) As Integer

Private Declare Function Sistema_ExecutarRelInt Lib "ADCUSR.DLL" Alias "AD_Sistema_ExecutarRel" (ByVal lpSist As Long, ByVal lpNomeTsk As String, ByVal lpRelOp As Long, ByVal Previa As Integer, ByVal dtDataAtual As Date, ByVal hWnd As Long, ByVal iOutOfProcess As Integer) As Long
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function Sist_ImpressoraDlgInt Lib "ADCUSR.DLL" Alias "AD_Sistema_PrinterDlg" (ByVal lID_Sistema As Long, ByVal bConfig As Integer, ByVal iLandscape As Integer) As Long

Private Declare Function Sistema_EditarRelInt Lib "ADCUSR.DLL" Alias "AD_Sistema_EditarRel" (ByVal lpSist As Long, ByVal lpNomeTsk As String) As Long
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function Sist_DefEmpresaInt Lib "ADCUSR.DLL" Alias "AD_Acesso_DefEmpresa" (ByVal lID_Acesso As Long, ByVal lpNomeEmpresa As String, ByVal iCodFilial As Integer, ByVal lpNomeFilial As String) As Long
'Define a Empresa corrente
'Retorno:
'    AD_BOOL_TRUE p/sucesso

Private Declare Function Sist_LoginInt Lib "ADCUSR.DLL" Alias "AD_Acesso_Login" (ByVal lID_Acesso As Long, ByVal lpCodUsuario As String, ByVal lpSenha As String) As Long
' Executa o login de um usuario
'Retorno:
'    AD_BOOL_TRUE p/sucesso

Private Declare Function Obter_GrupoInt Lib "ADCUSR.DLL" Alias "AD_Acesso_ObterGrupo" (ByVal lID_Acesso As Long, ByVal sCodGrupo As String) As Long
Private Declare Function Obter_UsuarioInt Lib "ADCUSR.DLL" Alias "AD_Acesso_ObterUsuario" (ByVal lID_Acesso As Long, ByVal sUsuario As String) As Long

Private Declare Function Rotina_ObterFuncaoInt Lib "ADCUSR.DLL" Alias "AD_DicDados_ObterFunRotVar" (ByVal IID_DicDados As Long, ByVal lpCodigo As String, lpVar As Variant) As Long
Private Declare Function Tela_ObterFuncaoInt Lib "ADCUSR.DLL" Alias "AD_DicDados_ObterFunTela" (ByVal IID_DicDados As Long, ByVal lpCodigo As String, ByVal lpProjeto As String, ByVal lpClasse As String) As Long
'Private Declare Function MenuItem_ObterFuncaoInt Lib "ADCUSR.DLL" Alias "AD_DicDados_ObterFunMenuItem" (ByVal IID_DicDados As Long, ByVal lpCodigo As String, ByVal lpProjeto As String, ByVal lpClasse As String) As Long

'para Sort
Private Declare Function Sort_AddKey Lib "ADCRTL.DLL" Alias "FN_Sort_AddKey" (ByVal lID_Sort As Long, lPosicion As Long, sCampo As Variant) As Long
Private Declare Function Sort_AddKeySeg Lib "ADCRTL.DLL" Alias "FN_Sort_AddKeySeg" (ByVal lID_Sort As Long, sCampo As Variant) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private mvarColObjFunc As Object
Private mvarvBufferObterRot As Variant
'Para o Loja (Col de Rotinas)
Private mvarcolRotinasECF As Object
Public gcolRotinasECF As New Collection '??? LIXO, vai ser apagado

Public Property Let gbTelaReordenando(vData As Boolean)
    SistGlob_bTelaReordenando = vData
End Property

Public Property Get gbTelaReordenando() As Boolean
    gbTelaReordenando = SistGlob_bTelaReordenando
End Property

Public Property Let glADMCount(ByVal vData As Long)
    SistGlob_lADMCount = vData
End Property

Public Property Get glADMCount() As Long
    glADMCount = SistGlob_lADMCount
End Property

Public Property Get gcolRotinasSGE() As Object
    Set gcolRotinasSGE = SistGlob_colRotinasSGE
End Property

Public Property Set gcolRotinasSGE(ByVal vData As Object)
    Set SistGlob_colRotinasSGE = vData
End Property

Public Property Let gbVPN(vData As Boolean)
    SistGlob_bVPN = vData
End Property

Public Property Get gbVPN() As Boolean
    gbVPN = SistGlob_bVPN
End Property

Public Property Let gbPreLoadGravar(vData As Boolean)
    SistGlob_bPreLoadGravar = vData
End Property

Public Property Get gbPreLoadGravar() As Boolean
    gbPreLoadGravar = SistGlob_bPreLoadGravar
End Property

Public Property Let giLocalOperacao(vData As Integer)
    SistGlob_iLocalOperacao = vData
End Property

Public Property Get giLocalOperacao() As Integer

    If SistGlob_iLocalOperacao = 0 Then
    
       SistGlob_iLocalOperacao = GetPrivateProfileInt("Geral", "LocalOperacao", LOCALOPERACAO_BACKOFFICE, "ADM100.INI")
    
    End If

    giLocalOperacao = SistGlob_iLocalOperacao
    
End Property

Public Property Let giContabGerencial(vData As Integer)
    SistGlob_iContabGerencial = vData
End Property

Public Property Get giContabGerencial() As Integer

    giContabGerencial = SistGlob_iContabGerencial
    
End Property

Public Property Let giFilialAuxiliar(vData As Integer)
    SistGlob_iFilialAuxiliar = vData
End Property

Public Property Get giFilialAuxiliar() As Integer

    giFilialAuxiliar = SistGlob_iFilialAuxiliar
    
End Property


Public Property Get CRLF() As String
    
    CRLF = Chr(13) & Chr(10)

End Property

Public Property Get gcolTiposMovEst() As Collection
    Set gcolTiposMovEst = SistGlob_colTiposMovEst
End Property

Public Property Set gcolTiposMovEst(ByVal vData As Collection)
    Set SistGlob_colTiposMovEst = vData
End Property

Public Property Get gcolErrosBatch() As Collection
    Set gcolErrosBatch = SistGlob_colErrosBatch
End Property

Public Property Set gcolErrosBatch(ByVal vData As Collection)
    Set SistGlob_colErrosBatch = vData
End Property

Public Property Get gcolCampos() As Collection
    Set gcolCampos = SistGlob_colCampos
End Property

Public Property Set gcolCampos(ByVal vData As Collection)
    Set SistGlob_colCampos = vData
End Property

Public Property Get gcolModulo() As AdmColModulo
    Set gcolModulo = SistGlob_objAdmColModulo
End Property

Public Property Set gcolModulo(ByVal vData As AdmColModulo)
    Set SistGlob_objAdmColModulo = vData
End Property

Public Property Set gobjButton(ByVal vData As Picture)
    Set SistGlob_objButton = vData
End Property

Public Property Get gobjButton() As Picture
    Set gobjButton = SistGlob_objButton
End Property

Public Property Set gobjOptionButtonUnChecked(ByVal vData As Picture)
    Set SistGlob_objOptionButtonUnchecked = vData
End Property

Public Property Get gobjOptionButtonUnChecked() As Picture
    Set gobjOptionButtonUnChecked = SistGlob_objOptionButtonUnchecked
End Property

Public Property Set gobjOptionButtonChecked(ByVal vData As Picture)
    Set SistGlob_objOptionButtonChecked = vData
End Property

Public Property Get gobjOptionButtonChecked() As Picture
    Set gobjOptionButtonChecked = SistGlob_objOptionButtonChecked
End Property

Public Property Set gobjCheckboxUnchecked(ByVal vData As Picture)
    Set SistGlob_objCheckboxUnchecked = vData
End Property

Public Property Get gobjCheckboxUnchecked() As Picture
    Set gobjCheckboxUnchecked = SistGlob_objCheckboxUnchecked
End Property

Public Property Set gobjCheckboxChecked(ByVal vData As Picture)
    Set SistGlob_objCheckboxChecked = vData
End Property

Public Property Get gobjCheckboxChecked() As Picture
    Set gobjCheckboxChecked = SistGlob_objCheckboxChecked
End Property

Public Property Set gobjCheckboxGrayed(ByVal vData As Picture)
    Set SistGlob_objCheckboxGrayed = vData
End Property

Public Property Get gobjCheckboxGrayed() As Picture
    Set gobjCheckboxGrayed = SistGlob_objCheckboxGrayed
End Property

Public Property Let gdtDataAtual(ByVal vData As Date)
    SistGlob_dtDataAtual = vData
End Property

Public Property Get gdtDataAtual() As Date
    gdtDataAtual = SistGlob_dtDataAtual
End Property

Public Property Let gdtDataHoje(ByVal vData As Date)
    SistGlob_dtDataHoje = vData
End Property

Public Property Get gdtDataHoje() As Date
    gdtDataHoje = SistGlob_dtDataHoje
End Property

Public Property Let GL_lUltimoErro(ByVal vData As Long)
    SistGlob_lUltimoErro = vData
End Property

Public Property Let giDebug(ByVal vData As Integer)
    SistGlob_iDebug = vData
End Property

Public Property Get giDebug() As Integer
    giDebug = SistGlob_iDebug
End Property

Public Property Let GL_lSistema(ByVal vData As Long)
    SistGlob_lSistema = vData
End Property

Public Property Let GL_lAcesso(ByVal vData As Long)
    SistGlob_lAcesso = vData
End Property

Public Property Let GL_lConexao(ByVal vData As Long)
    SistGlob_lConexao = vData
End Property

Public Property Let GL_lConexaoDic(ByVal vData As Long)
    SistGlob_lConexaoDic = vData
End Property

Public Property Let GL_lConexaoRel(ByVal vData As Long)
    SistGlob_lConexaoRel = vData
End Property

Public Property Let GL_lConexaoBrowse(ByVal vData As Long)
    SistGlob_lConexaoBrowse = vData
End Property

Public Property Let GL_lConexaoDicBrowse(ByVal vData As Long)
    SistGlob_lConexaoDicBrowse = vData
End Property

Public Property Let GL_lTransacao(ByVal vData As Long)
    SistGlob_lTransacao = vData
End Property

Public Property Let GL_lMascara(ByVal vData As Long)
    SistGlob_lMascara = vData
End Property

Public Property Get GL_objMascaraGenerica() As Object
    Set GL_objMascaraGenerica = SistGlob_objMascaraGenerica
End Property

Public Property Set GL_objMascaraGenerica(ByVal vData As Object)
    Set SistGlob_objMascaraGenerica = vData
End Property

Public Property Let GL_lMascProd(ByVal vData As Long)
    SistGlob_lMascProd = vData
End Property

Public Property Let GL_lDicDados(ByVal vData As Long)
    SistGlob_lDicDados = vData
End Property

Public Property Let glEmpresa(ByVal vData As Long)
    SistGlob_lEmpresa = vData
End Property

Public Property Let gsNomeEmpresa(vData As String)
    SistGlob_sNomeEmpresa = vData
End Property

Public Property Let gsNomePrinc(vData As String)
    SistGlob_sNomePrinc = vData
End Property

Public Property Let gsUsuario(vData As String)
    SistGlob_sUsuario = vData
End Property

Public Property Let giFilialEmpresa(ByVal vData As Integer)
    SistGlob_iFilialEmpresa = vData
End Property

Public Property Let giGrupoEmpresarial(ByVal vData As Integer)
    SistGlob_iGrupoEmpresarial = vData
End Property

Public Property Let giTelaTamanhoVariavel(ByVal vData As Integer)
    SistGlob_iTelaTamanhoVariavel = vData
End Property

Public Property Get giTelaTamanhoVariavel() As Integer
    giTelaTamanhoVariavel = SistGlob_iTelaTamanhoVariavel
End Property

Public Property Let gsNomeFilialEmpresa(vData As String)
    SistGlob_sNomeFilialEmpresa = vData
End Property

Public Property Set GL_objContabInt(ByVal vData As Object)
    Set SistGlob_objContabInt = vData
End Property

Public Property Set GL_objKeepAlive(ByVal vData As Object)
    Set SistGlob_objKeepAlive = vData
End Property

Public Property Set GL_objMDIForm(ByVal vData As Object)
    Set SistGlob_objMDIForm = vData
End Property

Public Property Get GL_lUltimoErro() As Long
    GL_lUltimoErro = SistGlob_lUltimoErro
End Property

Public Property Get GL_lSistema() As Long
    GL_lSistema = SistGlob_lSistema
End Property

Public Property Get GL_lAcesso() As Long
    GL_lAcesso = SistGlob_lAcesso
End Property

Public Property Get GL_lConexao() As Long
    GL_lConexao = SistGlob_lConexao
End Property

Public Property Get GL_lConexaoDic() As Long
    GL_lConexaoDic = SistGlob_lConexaoDic
End Property

Public Property Get GL_lConexaoRel() As Long
    GL_lConexaoRel = SistGlob_lConexaoRel
End Property

Public Property Get GL_lConexaoBrowse() As Long
    GL_lConexaoBrowse = SistGlob_lConexaoBrowse
End Property

Public Property Get GL_lConexaoDicBrowse() As Long
    GL_lConexaoDicBrowse = SistGlob_lConexaoDicBrowse
End Property

Public Property Get GL_lTransacao() As Long
    GL_lTransacao = SistGlob_lTransacao
End Property

Public Property Get GL_lTransacaoDic() As Long
    GL_lTransacaoDic = SistGlob_lTransacaoDic
End Property

Public Property Get GL_lMascara() As Long
    GL_lMascara = SistGlob_lMascara
End Property

Public Property Get GL_lMascProd() As Long
    GL_lMascProd = SistGlob_lMascProd
End Property

Public Property Get GL_lDicDados() As Long
    GL_lDicDados = SistGlob_lDicDados
End Property

Public Property Get glEmpresa() As Long
    glEmpresa = SistGlob_lEmpresa
End Property

Public Property Get gsNomeEmpresa() As String
    gsNomeEmpresa = SistGlob_sNomeEmpresa
End Property

Public Property Get gsNomePrinc() As String
    gsNomePrinc = SistGlob_sNomePrinc
End Property

Public Property Get gsUsuario() As String
    gsUsuario = SistGlob_sUsuario
End Property

Public Property Get giGrupoEmpresarial() As Integer
    giGrupoEmpresarial = SistGlob_iGrupoEmpresarial
End Property

Public Property Get giFilialEmpresa() As Integer
    giFilialEmpresa = SistGlob_iFilialEmpresa
End Property

Public Property Get gsNomeFilialEmpresa() As String
    gsNomeFilialEmpresa = SistGlob_sNomeFilialEmpresa
End Property

Public Property Let giDesconsideraFechamentoPeriodo(ByVal vData As Integer)
    SistGlob_iDesconsideraFechamentoPeriodo = vData
End Property

Public Property Get giDesconsideraFechamentoPeriodo() As Integer
    giDesconsideraFechamentoPeriodo = SistGlob_iDesconsideraFechamentoPeriodo
End Property

Public Property Let gsModulo(ByVal vData As String)
    SistGlob_sModulo = vData
End Property

Public Property Get gsModulo() As String
    gsModulo = SistGlob_sModulo
End Property

Public Property Get GL_objContabInt() As Object
    Set GL_objContabInt = SistGlob_objContabInt
End Property

Public Property Get GL_objKeepAlive() As AdmKeepAlive
    Set GL_objKeepAlive = SistGlob_objKeepAlive
End Property

Public Property Get GL_objMDIForm() As Object
    Set GL_objMDIForm = SistGlob_objMDIForm
End Property
Public Property Let giTipoVersao(ByVal vData As Integer)
    SistGlob_iTipoVersao = vData
End Property
Public Property Get giTipoVersao() As Integer
    giTipoVersao = SistGlob_iTipoVersao
End Property

Public Property Get gobjEstInicial() As Object
    Set gobjEstInicial = SistGlob_objEstInicial
End Property

Public Property Set gobjEstInicial(ByVal vData As Object)
    Set SistGlob_objEstInicial = vData
End Property

Public Property Get glpPrevWndProc() As Long
    glpPrevWndProc = SistGlob_lpPrevWndProc
End Property

Public Property Let glpPrevWndProc(ByVal vData As Long)
    SistGlob_lpPrevWndProc = vData
End Property

Public Property Get glpPrevWndProc0() As Long
    glpPrevWndProc0 = SistGlob_lpPrevWndProc0
End Property

Public Property Let glpPrevWndProc0(ByVal vData As Long)
    SistGlob_lpPrevWndProc0 = vData
End Property

Public Property Get glpPrevWndProc00() As Long
    glpPrevWndProc00 = SistGlob_lpPrevWndProc00
End Property

Public Property Let glpPrevWndProc00(ByVal vData As Long)
    SistGlob_lpPrevWndProc00 = vData
End Property

Public Property Get glpPrevWndProc1() As Long
    glpPrevWndProc1 = SistGlob_lpPrevWndProc1
End Property

Public Property Let glpPrevWndProc1(ByVal vData As Long)
    SistGlob_lpPrevWndProc1 = vData
End Property

Public Property Get glpPrevWndProc2() As Long
    glpPrevWndProc2 = SistGlob_lpPrevWndProc2
End Property

Public Property Let glpPrevWndProc2(ByVal vData As Long)
    SistGlob_lpPrevWndProc2 = vData
End Property

Public Property Get glpPrevWndProc3() As Long
    glpPrevWndProc3 = SistGlob_lpPrevWndProc3
End Property

Public Property Let glpPrevWndProc3(ByVal vData As Long)
    SistGlob_lpPrevWndProc3 = vData
End Property

Public Property Get glpPrevWndProc4() As Long
    glpPrevWndProc4 = SistGlob_lpPrevWndProc4
End Property

Public Property Let glpPrevWndProc4(ByVal vData As Long)
    SistGlob_lpPrevWndProc4 = vData
End Property

Public Property Get glpPrevWndProc5() As Long
    glpPrevWndProc5 = SistGlob_lpPrevWndProc5
End Property

Public Property Let glpPrevWndProc5(ByVal vData As Long)
    SistGlob_lpPrevWndProc5 = vData
End Property

Public Property Get glpPrevWndProc6() As Long
    glpPrevWndProc6 = SistGlob_lpPrevWndProc6
End Property

Public Property Let glpPrevWndProc6(ByVal vData As Long)
    SistGlob_lpPrevWndProc6 = vData
End Property

Public Property Get glpPrevWndProc7() As Long
    glpPrevWndProc7 = SistGlob_lpPrevWndProc7
End Property

Public Property Let glpPrevWndProc7(ByVal vData As Long)
    SistGlob_lpPrevWndProc7 = vData
End Property

Public Property Get glpPrevWndProc8() As Long
    glpPrevWndProc8 = SistGlob_lpPrevWndProc8
End Property

Public Property Let glpPrevWndProc8(ByVal vData As Long)
    SistGlob_lpPrevWndProc8 = vData
End Property

Public Property Get glpPrevWndProc9() As Long
    glpPrevWndProc9 = SistGlob_lpPrevWndProc9
End Property

Public Property Let glpPrevWndProc9(ByVal vData As Long)
    SistGlob_lpPrevWndProc9 = vData
End Property

Public Property Get glpPrevWndProc10() As Long
    glpPrevWndProc10 = SistGlob_lpPrevWndProc10
End Property

Public Property Let glpPrevWndProc10(ByVal vData As Long)
    SistGlob_lpPrevWndProc10 = vData
End Property

Public Property Get glpPrevWndProc11() As Long
    glpPrevWndProc11 = SistGlob_lpPrevWndProc11
End Property

Public Property Let glpPrevWndProc11(ByVal vData As Long)
    SistGlob_lpPrevWndProc11 = vData
End Property

Public Property Get glpPrevWndProc12() As Long
    glpPrevWndProc12 = SistGlob_lpPrevWndProc12
End Property

Public Property Let glpPrevWndProc12(ByVal vData As Long)
    SistGlob_lpPrevWndProc12 = vData
End Property

Public Property Get glpPrevWndProc13() As Long
    glpPrevWndProc13 = SistGlob_lpPrevWndProc13
End Property

Public Property Let glpPrevWndProc13(ByVal vData As Long)
    SistGlob_lpPrevWndProc13 = vData
End Property

Public Property Get glpPrevWndProc14() As Long
    glpPrevWndProc14 = SistGlob_lpPrevWndProc14
End Property

Public Property Let glpPrevWndProc14(ByVal vData As Long)
    SistGlob_lpPrevWndProc14 = vData
End Property

Public Property Get glpPrevWndProc15() As Long
    glpPrevWndProc15 = SistGlob_lpPrevWndProc15
End Property

Public Property Let glpPrevWndProc15(ByVal vData As Long)
    SistGlob_lpPrevWndProc15 = vData
End Property

Public Property Get gobjTelaAtiva() As Object
    Set gobjTelaAtiva = SistGlob_objTelaAtiva
End Property

Public Property Set gobjTelaAtiva(ByVal vData As Object)
    Set SistGlob_objTelaAtiva = vData
End Property

Public Property Get gobjControleDrag() As Object
    Set gobjControleDrag = SistGlob_objControleDrag
End Property

Public Property Set gobjControleDrag(ByVal vData As Object)
    Set SistGlob_objControleDrag = vData
End Property

Public Property Get gobjControleAlvo() As Object
    Set gobjControleAlvo = SistGlob_objControleAlvo
End Property

Public Property Set gobjControleAlvo(ByVal vData As Object)
    Set SistGlob_objControleAlvo = vData
End Property

Public Property Get gobjmenuEdicao() As Object
    Set gobjmenuEdicao = SistGlob_objmenuEdicao
End Property

Public Property Set gobjmenuEdicao(ByVal vData As Object)
    Set SistGlob_objmenuEdicao = vData
End Property

Public Property Get gcolWndProc() As Collection
    Set gcolWndProc = SistGlob_colWndProc
End Property

Public Property Set gcolWndProc(ByVal vData As Collection)
    Set SistGlob_colWndProc = vData
End Property

Public Property Get giLeft() As Integer
    giLeft = SistGlob_iLeft
End Property

Public Property Let giLeft(ByVal vData As Integer)
    SistGlob_iLeft = vData
End Property

Public Property Get giTop() As Integer
    giTop = SistGlob_iTop
End Property

Public Property Let giTop(ByVal vData As Integer)
    SistGlob_iTop = vData
End Property

Public Property Get gsngEdicaoX() As Single
    gsngEdicaoX = SistGlob_sngEdicaoX
End Property

Public Property Let gsngEdicaoX(ByVal vData As Single)
    SistGlob_sngEdicaoX = vData
End Property

Public Property Get gsngEdicaoY() As Single
    gsngEdicaoY = SistGlob_sngEdicaoY
End Property

Public Property Let gsngEdicaoY(ByVal vData As Single)
    SistGlob_sngEdicaoY = vData
End Property

Public Property Get gsngDragX() As Single
    gsngDragX = SistGlob_sngDragX
End Property

Public Property Let gsngDragX(ByVal vData As Single)
    SistGlob_sngDragX = vData
End Property

Public Property Get gsngDragY() As Single
    gsngDragY = SistGlob_sngDragY
End Property

Public Property Let gsngDragY(ByVal vData As Single)
    SistGlob_sngDragY = vData
End Property

Public Property Get gobjPropriedades() As Object
    Set gobjPropriedades = SistGlob_objPropriedades
End Property

Public Property Set gobjPropriedades(ByVal vData As Object)
    Set SistGlob_objPropriedades = vData
End Property

Public Property Get gobjCamposInvisiveis() As Object
    Set gobjCamposInvisiveis = SistGlob_objCamposInvisiveis
End Property

Public Property Set gobjCamposInvisiveis(ByVal vData As Object)
    Set SistGlob_objCamposInvisiveis = vData
End Property

Public Property Get giProxMouseMove() As Integer
    giProxMouseMove = SistGlob_iProxMouseMove
End Property

Public Property Let giProxMouseMove(ByVal vData As Integer)
    SistGlob_iProxMouseMove = vData
End Property

Public Property Get giProxButtonUp() As Integer
    giProxButtonUp = SistGlob_iProxButtonUp
End Property

Public Property Let giProxButtonUp(ByVal vData As Integer)
    SistGlob_iProxButtonUp = vData
End Property

'Edicao Tela - Raphael
Public Property Get gcolEdicaoTela() As Collection
    Set gcolEdicaoTela = SistGlob_colEdicaoTela
End Property

Public Property Set gcolEdicaoTela(ByVal vData As Collection)
    Set SistGlob_colEdicaoTela = vData
End Property

Public Property Get giBrowsePosicaoAntigo() As Integer
    giBrowsePosicaoAntigo = SistGlob_iBrowsePosicaoAntigo
End Property

Public Property Let giBrowsePosicaoAntigo(ByVal vData As Integer)
    SistGlob_iBrowsePosicaoAntigo = vData
End Property

Public Property Get giCliAtrasoDestacar() As Integer
    giCliAtrasoDestacar = SistGlob_iCliAtrasoDestacar
End Property

Public Property Let giCliAtrasoDestacar(ByVal vData As Integer)
    SistGlob_iCliAtrasoDestacar = vData
End Property

Public Property Get gsExtensaoGerRelExp() As String
    gsExtensaoGerRelExp = SistGlob_sExtensaoGerRelExp
End Property

Public Property Let gsExtensaoGerRelExp(ByVal vData As String)
    SistGlob_sExtensaoGerRelExp = vData
End Property

Function Sort_Inserir(ByVal lID_Sort As Long, ByVal lPosicao As Long, vSegmento1 As Variant, Optional vSegmento2 As Variant, Optional vSegmento3 As Variant, Optional vSegmento4 As Variant, Optional vSegmento5 As Variant) As Long
'**********************************************************************
'**********************************************************************
'**********************************************************************
'**********************************************************************
On Error GoTo saida_com_erro

If (Sort_AddKey(lID_Sort, lPosicao, vSegmento1) <> AD_BOOL_TRUE) Then Error 1

If (IsMissing(vSegmento2)) Then GoTo saida_sucesso
If (Sort_AddKeySeg(lID_Sort, vSegmento2) <> AD_BOOL_TRUE) Then Error 1

If (IsMissing(vSegmento3)) Then GoTo saida_sucesso
If (Sort_AddKeySeg(lID_Sort, vSegmento3) <> AD_BOOL_TRUE) Then Error 1

If (IsMissing(vSegmento4)) Then GoTo saida_sucesso
If (Sort_AddKeySeg(lID_Sort, vSegmento4) <> AD_BOOL_TRUE) Then Error 1

If (IsMissing(vSegmento5)) Then GoTo saida_sucesso
If (Sort_AddKeySeg(lID_Sort, vSegmento5) <> AD_BOOL_TRUE) Then Error 1

saida_sucesso:
Sort_Inserir = AD_BOOL_TRUE
Exit Function

saida_com_erro:
Sort_Inserir = AD_BOOL_FALSE
Exit Function
End Function

Function Sistema_QueryUnload() As Boolean

    If (GL_lSistema = 0) Then
        Sistema_QueryUnload = True
    Else
        Sistema_QueryUnload = Sistema_QueryUnloadInt(GL_lSistema)
    End If
    
End Function

Function Sistema_Abrir() As Long
'**********************************************************************
'**********************************************************************
'    iniciar utilizacao do Sistema
'Retorno:
'    NULL p/erro
'**********************************************************************
'**********************************************************************

Dim X As New AdmSQL, Y As New ClassConstCust

On Error GoTo saida_com_erro
   
    'Call Y.Inicializa_Tamanhos_String
    
    GL_lUltimoErro = 0
    
    GL_lSistema = Sistema_CriarInt()
    If (GL_lSistema = 0) Then Error 1
    
    GL_lAcesso = Sist_ObterAcesso(GL_lSistema)
    GL_lDicDados = Sist_ObterDicDados(GL_lSistema)
    GL_lConexaoDic = Sist_ObterConDic(GL_lSistema)
    GL_lConexaoDicBrowse = Sist_ObterConDicBrowse(GL_lSistema)
    Call X.Conexao_ObterOrdParam(GL_lConexaoDic, SistGlob_iSQLTipoOrdParamDic)
    
    giTipoVersao = Sist_ObterTipoVersao(GL_lSistema)
            
    Sistema_Abrir = GL_lSistema
    
    Exit Function

saida_com_erro:

    Sistema_Abrir = 0
    
    MsgBox (Error)
    
    Exit Function
    
End Function


Function Sistema_Preparar_Batch(sNomeArqParam As String) As Long
'cria um arquivo com dados a serem utilizados na ativacao de um batch
Dim lErro As Long
On Error GoTo Erro_Sistema_Preparar_Batch

    sNomeArqParam = String(200, 0)
    
    lErro = Sistema_ExecutarBatch_Int(GL_lSistema, sNomeArqParam, giFilialEmpresa, gsNomeFilialEmpresa, gdtDataAtual)
    If lErro <> AD_BOOL_TRUE Then Error 27438
             
    sNomeArqParam = StringZ(sNomeArqParam)
    
    Sistema_Preparar_Batch = SUCESSO

    Exit Function

Erro_Sistema_Preparar_Batch:

    Sistema_Preparar_Batch = Err

    Select Case Err
    
        Case 27438
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_BATCH", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142005)

    End Select

    Exit Function

End Function

Function Sistema_Abrir_Batch(sNomeArqParam As String) As Long
'**********************************************************************
'**********************************************************************
'    iniciar utilizacao do Sistema para um batch
'    sNomeArqParam é o nome do arquivo criado p/passar ao batch dados do aplicativo chamador
'
'**********************************************************************
'**********************************************************************
Dim lCodEmp As Long, sNomeEmpresa As String, iCodFilial As Integer, sNomeFilial As String, dtDataAtual As Date, sUsu As String
Dim X As New AdmSQL, Y As New ClassConstCust

On Error GoTo saida_com_erro
   
    'Call Y.Inicializa_Tamanhos_String
    
    gbVPN = False
    
    GL_lUltimoErro = 0
    
    sNomeFilial = String(255, 0)
    sNomeEmpresa = String(255, 0)
    
    GL_lSistema = Sistema_Criar_Batch_Int(sNomeArqParam, lCodEmp, sNomeEmpresa, iCodFilial, sNomeFilial, dtDataAtual)
    If (GL_lSistema = 0) Then Error 1
    
    GL_lAcesso = Sist_ObterAcesso(GL_lSistema)
    If (GL_lAcesso = 0) Then Error 2
            
    GL_lConexaoDic = Sist_ObterConDic(GL_lSistema)
    If (GL_lConexaoDic = 0) Then Error 7500
    Call X.Conexao_ObterOrdParam(GL_lConexaoDic, SistGlob_iSQLTipoOrdParamDic)
    
    GL_lConexao = Sist_ObterConexao(GL_lSistema)
    Call X.Conexao_ObterOrdParam(GL_lConexao, SistGlob_iSQLTipoOrdParam)
    
    GL_lDicDados = Sist_ObterDicDados(GL_lSistema)

    giTipoVersao = Sist_ObterTipoVersao(GL_lSistema)
    
    glEmpresa = lCodEmp
    sNomeEmpresa = StringZ(sNomeEmpresa)
    gsNomeEmpresa = sNomeEmpresa
    
    giFilialEmpresa = iCodFilial
    sNomeFilial = StringZ(sNomeFilial)
    gsNomeFilialEmpresa = sNomeFilial
       
    sUsu = String(255, 0)
    Call Obter_UsuarioInt(GL_lAcesso, sUsu)
    sUsu = StringZ(sUsu)
    gsUsuario = sUsu
    
    gdtDataAtual = dtDataAtual
    gdtDataHoje = dtDataAtual
    
    If Sistema_Inicializa_Modulos <> SUCESSO Then Error 41629
    
    Call Y.Inicializa_Tamanhos_String
    
    Sistema_Abrir_Batch = SUCESSO

    Exit Function

saida_com_erro:

    Sistema_Abrir_Batch = 0
    
    Exit Function
    
End Function

Function Sistema_Abrir_Rel(lSistema As Long, lCodEmp As Long, sNomeEmpresa As String, iCodFilial As Integer, sNomeFilial As String, dtDataAtual As Date) As Long
'**********************************************************************
'**********************************************************************
'    iniciar utilizacao do Sistema para auxiliar a execucao de um relatorio outofprocess (execucao inprocess nao passa por aqui)
' ATENCAO: SOMENTE p/EXECUCAO OUTOFPROCESS (via fwexe4.exe)
'**********************************************************************
'**********************************************************************

Dim sUsu As String
Dim X As New AdmSQL, Y As New ClassConstCust

On Error GoTo saida_com_erro
   
    'Call Y.Inicializa_Tamanhos_String
    
    gbVPN = False
    
    GL_lUltimoErro = 0
    
    GL_lSistema = lSistema
    If (GL_lSistema = 0) Then Error 1
    
    GL_lAcesso = Sist_ObterAcesso(GL_lSistema)
    If (GL_lAcesso = 0) Then Error 2
            
    GL_lConexaoDic = Sist_ObterConDic(GL_lSistema)
    If (GL_lConexaoDic = 0) Then Error 7500
    Call X.Conexao_ObterOrdParam(GL_lConexaoDic, SistGlob_iSQLTipoOrdParamDic)

    GL_lConexao = Sist_ObterConexao(GL_lSistema)
    Call X.Conexao_ObterOrdParam(GL_lConexao, SistGlob_iSQLTipoOrdParam)
    
    GL_lConexaoRel = GL_lConexao
    
    GL_lDicDados = Sist_ObterDicDados(GL_lSistema)

    giTipoVersao = Sist_ObterTipoVersao(GL_lSistema)
    
    glEmpresa = lCodEmp
    gsNomeEmpresa = sNomeEmpresa
    
    giFilialEmpresa = iCodFilial
    gsNomeFilialEmpresa = sNomeFilial
    
    gdtDataAtual = dtDataAtual
    gdtDataHoje = dtDataAtual
       
    sUsu = String(255, 0)
    Call Obter_UsuarioInt(GL_lAcesso, sUsu)
    sUsu = StringZ(sUsu)
    gsUsuario = sUsu
    
    If Sistema_Inicializa_Modulos <> SUCESSO Then Error 41630
    
    Call Y.Inicializa_Tamanhos_String
    
    Sistema_Abrir_Rel = SUCESSO

    Exit Function

saida_com_erro:

    Sistema_Abrir_Rel = 0
    
    Exit Function
    
End Function

Function Sistema_Fechar() As Long
'**********************************************************************
'**********************************************************************
'    iniciar utilizacao do Sistema
'Retorno:
'    AD_BOOL_TRUE p/sucesso
'**********************************************************************
'**********************************************************************
    
    Call Sistema_Libera_Globais
    
    Set GL_objContabInt = Nothing
    Set mvarColObjFunc = Nothing
    Set gcolRotinasSGE = Nothing
    Set GL_objKeepAlive = Nothing
    
    Sistema_Fechar = Sistema_DestruirInt(GL_lSistema)
    GL_lSistema = 0
    GL_lAcesso = 0
    GL_lMascara = 0
    GL_lDicDados = 0
    GL_lConexaoDic = 0
    
End Function

Function Sistema_Inicializa_Modulos() As Long
'??? pendencia: deveria testar se os modulos estao ativos mas como as classes estao em vbps diferentes vou adiar este teste. jones 22/03/99
Dim lErro As Long, objAux As Object
    
On Error GoTo Erro_Sistema_Inicializa_Modulos

    
    'para liberar os objetos globais que estavam carregados
    Set GL_objKeepAlive = New AdmKeepAlive
    
    'If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisContab.ClassContabInt")
        If objAux Is Nothing Then Error 41623
        Set GL_objContabInt = objAux
        GL_objKeepAlive.Add objAux
    'End If
    
        Set objAux = CreateObject("GlobaisCRFAT.ClassCRFATGlobal")
        If objAux Is Nothing Then Error 56653
        GL_objKeepAlive.Add objAux
    
        Set objAux = CreateObject("GlobaisMAT.ClassMATGlobal")
        If objAux Is Nothing Then Error 59268
        GL_objKeepAlive.Add objAux
    
    'If gcolModulo.Ativo(MODULO_CONTASAPAGAR) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisCPR.ClassCPGlobal")
        If objAux Is Nothing Then Error 41624
        GL_objKeepAlive.Add objAux
    'End If
    
    'If gcolModulo.Ativo(MODULO_CONTASARECEBER) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisCPR.ClassCRGlobal")
        If objAux Is Nothing Then Error 41625
        GL_objKeepAlive.Add objAux
    'End If
    
    'If gcolModulo.Ativo(MODULO_TESOURARIA) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisCPR.ClassTESGlobal")
        If objAux Is Nothing Then Error 41626
        GL_objKeepAlive.Add objAux
    'End If
    
    'If gcolModulo.Ativo(MODULO_ESTOQUE) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisEST.ClassESTGlobal")
        If objAux Is Nothing Then Error 41627
        GL_objKeepAlive.Add objAux
    'End If
    
    'If gcolModulo.Ativo(MODULO_FATURAMENTO) = MODULO_ATIVO Then
        Set objAux = CreateObject("GlobaisFAT.ClassFATGlobal")
        If objAux Is Nothing Then Error 41628
        GL_objKeepAlive.Add objAux
    'End If
    
    Sistema_Inicializa_Modulos = SUCESSO
     
    Exit Function
    
Erro_Sistema_Inicializa_Modulos:

    Sistema_Inicializa_Modulos = Err
     
    Select Case Err
          
        Case 41623 To 41628, 56653
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142006)
     
    End Select
     
    Exit Function

End Function

Function Sistema_Reseta_Modulos()

Dim lErro As Long, objAux As Object

On Error GoTo Erro_Sistema_Reseta_Modulos

    Set objAux = CreateObject("GlobaisCRFAT.ClassCRFATGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisMAT.ClassMATGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisCPR.ClassCPGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisCPR.ClassCRGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisCPR.ClassTESGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisEST.ClassESTGlobal")
    Call objAux.Resetar

    Set objAux = CreateObject("GlobaisFAT.ClassFATGlobal")
    Call objAux.Resetar
    
    Set objAux = CreateObject("GlobaisCOM.ClassCOMGlobal")
    Call objAux.Resetar
    
    Set objAux = CreateObject("GlobaisLoja.ClassLojaGlobal")
    Call objAux.Resetar
    
    Sistema_Reseta_Modulos = SUCESSO
     
    Exit Function
    
Erro_Sistema_Reseta_Modulos:

    Sistema_Reseta_Modulos = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142007)
     
    End Select
     
    Exit Function

End Function

Function Sistema_DefEmpresa(ByVal lpNomeEmpresa As String, ByVal lCodEmp As Long, ByVal lpNomeFilial As String, ByVal iCodFilial As Integer) As Long
'**********************************************************************
'**********************************************************************
' Define Empresa corrente
'Retorno:
'    AD_BOOL_TRUE p/sucesso
'**********************************************************************
'**********************************************************************
Dim ret As Integer, objAux As Object
Dim X As Object
Dim lErro As Long, Y As New AdmSQL

On Error GoTo saida_com_erro
    
    ret = Sist_DefEmpresaInt(GL_lAcesso, lpNomeEmpresa, iCodFilial, lpNomeFilial)
    If (ret <> AD_BOOL_TRUE) Then Error 1
    
    GL_lConexao = Sist_ObterConexao(GL_lSistema)
    Call Y.Conexao_ObterOrdParam(GL_lConexao, SistGlob_iSQLTipoOrdParam)
    
    GL_lConexaoBrowse = Sist_ObterConBrowse(GL_lSistema)
    GL_lConexaoRel = Sist_ObterConRel(GL_lSistema)
    
    glEmpresa = lCodEmp
    gsNomeEmpresa = lpNomeEmpresa
    giFilialEmpresa = iCodFilial
    gsNomeFilialEmpresa = lpNomeFilial
    
    lErro = CF("Retorna_ColFiliais")
    If lErro <> SUCESSO Then Error 55176
    
    If Sistema_Inicializa_Modulos <> SUCESSO Then Error 41622
    
    Sistema_DefEmpresa = ret
    
    Exit Function

saida_com_erro:
    
    Sistema_DefEmpresa = 0
    
    Exit Function

End Function

Function Sistema_Login(ByVal lpCodUsuario As String, ByVal lpSenha As String) As Long
'**********************************************************************
'**********************************************************************
' Define Empresa corrente
'Retorno:
'    AD_BOOL_TRUE p/sucesso
'**********************************************************************
'**********************************************************************
gsUsuario = lpCodUsuario

Sistema_Login = Sist_LoginInt(GL_lAcesso, lpCodUsuario, lpSenha)

End Function

Function Limpa_Tela(Frm As Object)

    Dim sTipo As String, obj2 As Object
    Dim objMasked As Object, objText As TextBox
    
    For Each obj2 In Frm.Controls

        sTipo = TypeName(obj2)

        If sTipo = "MaskEdBox" Then
            Set objMasked = obj2
            If objMasked.PromptInclude = True Then
                objMasked.PromptInclude = False
                objMasked.Text = ""
                objMasked.PromptInclude = True
            Else
                objMasked.Text = "" 'String$(.Controls(i).MaxLength, .Controls(i).PromptChar)
            End If
'        ElseIf sTipo = "SSPanel" Then
'            obj2.Caption = ""
        ElseIf sTipo = "TextBox" Then
            Set objText = obj2
            objText.Text = ""
        End If
    Next

End Function

Function Rotina_Erro(ByVal MsgBoxTipo As Integer, ByVal sErroId As String, ByVal lCodigo As Long, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As Long

Dim sErro As String, X As Object, Y As New AdmSQL
Dim iMsgBoxRet As Integer, lCodigoInt As Long, sTipoErro As String
Dim lErroAux As Integer, iTipoCliente As Integer
Dim lErro As Integer

On Error GoTo Erro_Rotina_Erro
    
    'Se o projeto é o SGEECF
    If gsNomePrinc = "SGEECF" Then
        Rotina_Erro = Rotina_ErroECF2(MsgBoxTipo, sErroId, lCodigo, vParam1, vParam2, vParam3, vParam4, vParam5, vParam6, vParam7, vParam8, vParam9, vParam10)
    Else

        iTipoCliente = Sistema_ObterTipoCliente(GL_lSistema)
        If iTipoCliente = AD_SIST_BATCH Then
            lCodigoInt = 0
        Else
            lCodigoInt = lCodigo
        End If
    
        'para depurar o batchest2 como uma dll o trecho abaixo até os asteriscos deve estar comentado
        'se há alguma transacao aberta vou forcar o rollback
        If GL_lTransacao <> 0 And iTipoCliente <> AD_SIST_BATCH Then
            Call Y.Transacao_Rollback
        End If
        
        If GL_lTransacaoDic <> 0 And iTipoCliente <> AD_SIST_BATCH Then
            Call Y.Transacao_RollbackExt(GL_lTransacaoDic)
        End If
    
    '****
        
        GL_lUltimoErro = lCodigo
    
        Rotina_Erro = 0
        sErro = String(1024, 0)
    
        Call Rotina_Erro_Carregar_Msg(sErroId, sTipoErro)
        
        'para depurar o batchest como uma dll o comando abaixo deve estar descomentado
    '    iTipoCliente = AD_SIST_BATCH
    '****
    
    
        If Not IsMissing(vParam10) Then
            lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam9) Then
                lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam8) Then
                    lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam7) Then
                        lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam6) Then
                            lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam5) Then
                                lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam4) Then
                                    lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam3) Then
                                        lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam2) Then
                                            lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        Else
                                            If Not IsMissing(vParam1) Then
                                                lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                                If lErroAux Then Exit Function
                                            Else
                                                lErroAux = Rotina_Erro_Carregar(sTipoErro, lCodigoInt, sErro, "", "", "", "", "", "", "", "", "", "")
                                                If lErroAux Then Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        Call Sistema_RegistrarOcorrencia(GL_lSistema, "Empresa: " & gsNomeEmpresa & " Filial: " & gsNomeFilialEmpresa & " - " & Format(Now(), "General Date") & " - " & "Usuário: " & gsUsuario)
        Call Sistema_RegistrarOcorrencia(GL_lSistema, "   ERRO: " & sTipoErro & " Local: " & CStr(lCodigo) & " Descrição: " & sErro)
        
        If iTipoCliente <> AD_SIST_BATCH Then
            
            Set X = CreateObject("TelasAdm.ClassTelasAdm")
            
            If X Is Nothing Then Error 58452
            If X.objFormMsgErro Is Nothing Then Error 58453
            
            X.objFormMsgErro.sErro = sErro
            X.objFormMsgErro.lLocalErro = lCodigo
            X.objFormMsgErro.sTipoErro = sTipoErro
            
            X.objFormMsgErro.Show vbModal
            'Rotina_Erro = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")
            
        Else
        
            If gcolErrosBatch Is Nothing Then Set gcolErrosBatch = New Collection
        
            gcolErrosBatch.Add "Local: " & CStr(lCodigo) & " Descrição: " & sErro
            
        End If
        
    End If
    Rotina_Erro = vbOK

    Exit Function
    
Erro_Rotina_Erro:
    
    Select Case Err
    
        Case 58452, 58453, 99730
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Exit Function

End Function

Function Rotina_ErrosBatch() As Long

Dim X As Object

On Error GoTo Erro_Rotina_ErrosBatch

    Set X = CreateObject("TelasAdm.ClassTelasAdm")
    
    If X Is Nothing Then Error 58452
    If X.objFormMsgErroBatch Is Nothing Then Error 58453
            
    X.objFormMsgErroBatch.Show vbModal
    
    Rotina_ErrosBatch = vbOK

    Exit Function
    
Erro_Rotina_ErrosBatch:
    
    Select Case Err
    
        Case 58452, 58453
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Exit Function
            
End Function

Function Rotina_ErrosBatch2(Optional ByVal sDescricao As String) As Long

Dim X As Object
Dim vValor As Variant
Dim iIndice As Integer
Dim sTexto As String

On Error GoTo Erro_Rotina_ErrosBatch2

    If Not (gcolErrosBatch Is Nothing) Then
        iIndice = 0
        sTexto = ""
        For Each vValor In gcolErrosBatch
            iIndice = iIndice + 1
            If Len(Trim(sTexto)) > 0 Then sTexto = sTexto & vbNewLine
            sTexto = sTexto & CStr(iIndice) & SEPARADOR & Replace(Replace(vValor, Chr(0), ""), "|", "")
        Next
        Set gcolErrosBatch = Nothing
    End If
    
    If sDescricao = "" Then sDescricao = " "
    
    Call WinExec(App.Path & "\ErrosBatch.exe " & sDescricao & "|" & sTexto & "|" & gsNomeEmpresa & "|" & gsNomeFilialEmpresa & "|" & gsUsuario, SW_NORMAL)
    
    Rotina_ErrosBatch2 = vbOK

    Exit Function
    
Erro_Rotina_ErrosBatch2:
    
    Select Case Err
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Exit Function
            
End Function

''Function Rotina_ObterFuncao(lpCodigo As String, lpProjeto As String, lpClasse As String) As Long
'''Retorno:
'''    AD_BOOL_TRUE p/sucesso
''    Rotina_ObterFuncao = Rotina_ObterFuncaoInt(GL_lDicDados, lpCodigo, lpProjeto, lpClasse)
''    lpProjeto = StringZ(lpProjeto)
''    lpClasse = StringZ(lpClasse)
''End Function
''
Function Tela_ObterFuncao(lpCodigo As String, lpProjeto As String, lpClasse As String) As Long
'Retorno:
'    AD_BOOL_TRUE p/sucesso
Dim sGrupoTelasProjeto, sGrupoTelasClasse, sTelasProjetoOriginal, sTelasClasseOriginal, sTelasProjetoCustomizado, sTelasClasseCustomizado
Dim lComando As Long, Y As New AdmSQL, sCodGrupo As String, iFilialEmpresa As Integer
Dim lErro As Long

On Error GoTo Erro_Tela_ObterFuncao

    If Not gbVPN Then
    
        lErro = Tela_ObterFuncaoInt(GL_lDicDados, lpCodigo, lpProjeto, lpClasse)
        lpProjeto = StringZ(lpProjeto)
        lpClasse = StringZ(lpClasse)

    Else
    
        Call Obter_Grupo(sCodGrupo)
        
        lComando = Y.Comando_AbrirExt(GL_lConexaoDicBrowse)
        If lComando = 0 Then Error 59277
            
        sGrupoTelasProjeto = String(NOME_PROJETO, 0)
        sGrupoTelasClasse = String(NOME_CLASSE, 0)
        sTelasProjetoOriginal = String(NOME_PROJETO, 0)
        sTelasClasseOriginal = String(NOME_CLASSE, 0)
        sTelasProjetoCustomizado = String(NOME_PROJETO, 0)
        sTelasClasseCustomizado = String(NOME_CLASSE, 0)
        
        lErro = Y.Comando_Executar(lComando, "SELECT GrupoTela.Projeto, GrupoTela.Classe, Telas.Projeto_original, Telas.Classe_original, Telas.Projeto_customizado, Telas.Classe_customizado, Telas.FilialEmpresa FROM GrupoTela, Telas WHERE GrupoTela.CodGrupo=? AND GrupoTela.NomeTela = Telas.Nome AND TipoDeAcesso = 1 AND Telas.Nome = ?", _
            sGrupoTelasProjeto, sGrupoTelasClasse, sTelasProjetoOriginal, sTelasClasseOriginal, sTelasProjetoCustomizado, sTelasClasseCustomizado, iFilialEmpresa, sCodGrupo, lpCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 59278
        
        lErro = Y.Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59279
        
        If lErro = AD_SQL_SUCESSO Then
            
            If sGrupoTelasProjeto = "" Then
                If sTelasProjetoCustomizado <> "" Then
                    lpProjeto = sTelasProjetoCustomizado
                    lpClasse = sTelasClasseCustomizado
                Else
                    lpProjeto = sTelasProjetoOriginal
                    lpClasse = sTelasClasseOriginal
                End If
            Else
                lpProjeto = sGrupoTelasProjeto
                lpClasse = sGrupoTelasClasse
            End If
            
            lErro = AD_BOOL_TRUE
            
        Else
        
            Error 59281
            
        End If
        
        Call Y.Comando_Fechar(lComando)
    
    End If
    
    Tela_ObterFuncao = lErro
    
    Exit Function
    
Erro_Tela_ObterFuncao:

    Tela_ObterFuncao = Err
    
    Select Case Err
          
        Case 59277
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 59278, 59279, 59280
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIC_ROTINAS", Err)
        
        Case 59281
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error & " Erro na leitura da tela " & lpCodigo, 142026)
     
    End Select
     
    If lComando <> 0 Then Call Y.Comando_Fechar(lComando)
    
    Exit Function
    
End Function

''Function MenuItem_ObterFuncao(lpCodigo As String, lpProjeto As String, lpClasse As String) As Long
'''Retorno:
'''    AD_BOOL_TRUE p/sucesso
''    MenuItem_ObterFuncao = MenuItem_ObterFuncaoInt(GL_lDicDados, lpCodigo, lpProjeto, lpClasse)
''End Function

Function Init_Contab() As Long
    Init_Contab = Init_Contab_Int(GL_lSistema)
    GL_lMascara = Sist_ObterMascara(GL_lSistema)
End Function

Function Reset_Contab() As Long
    Reset_Contab = Reset_Contab_Int(GL_lSistema)
    GL_lMascara = Sist_ObterMascara(GL_lSistema)
End Function

Function Init_Fest() As Long
    Init_Fest = Init_Fest_Int(GL_lSistema)
    GL_lMascProd = Sist_ObterMascaraProd(GL_lSistema)
End Function

Function Reset_Fest() As Long
    Reset_Fest = Reset_Fest_Int(GL_lSistema)
    GL_lMascProd = Sist_ObterMascaraProd(GL_lSistema)
End Function

Function StringZ(bstr As String) As String
    Dim iPos As Integer
    iPos = InStr(1, bstr, Chr$(0), 0)
    If iPos > 0 Then
        StringZ = Mid$(bstr, 1, iPos - 1)
    Else
        StringZ = bstr
    End If
End Function

Function Rotina_Aviso(ByVal MsgBoxTipo As VbMsgBoxStyle, sErroId As String, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As VbMsgBoxResult
''   Rotina_Aviso = MsgBox("Confirma a operação?", MsgBoxTipo, "Rotina temporaria de aviso")
Dim sErro As String
Dim lErroAux As Integer
Dim X As Object, sTipoErro As String
Dim asReg(1 To 100) As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Aviso
        
iIndice = 1

asReg(iIndice) = "****************************************************************************"
iIndice = iIndice + 1
asReg(iIndice) = Format(Now, "dd/mm/yyyy hh:mm:ss") & "    -    " & gsNomeEmpresa & "    -    " & gsUsuario
iIndice = iIndice + 1
asReg(iIndice) = "AdmLib.ADM.Rotina_Aviso "
iIndice = iIndice + 1
asReg(iIndice) = " "
iIndice = iIndice + 1

asReg(iIndice) = "1 - sErroId = " & sErroId
iIndice = iIndice + 1

asReg(iIndice) = "2 - gsNomePrinc = " & gsNomePrinc
iIndice = iIndice + 1

    'Se o projeto é o SGEECF
    If gsNomePrinc = "SGEECF" Then
    
asReg(iIndice) = "3 - Rotina_AvisoECF2"
iIndice = iIndice + 1

        Rotina_Aviso = Rotina_AvisoECF2(MsgBoxTipo, sErroId, vParam1, vParam2, vParam3, vParam4, vParam5, vParam6, vParam7, vParam8, vParam9, vParam10)
    
asReg(iIndice) = "4 - Rotina_Aviso (Rotina_AvisoECF2) = " & CStr(Rotina_Aviso)
iIndice = iIndice + 1
    
    Else
        Rotina_Aviso = 0
        sErro = String(1024, 0)
      
asReg(iIndice) = "5 - Rotina_Erro_Carregar_Msg. sErroId  = " & sErroId
iIndice = iIndice + 1

        Call Rotina_Erro_Carregar_Msg(sErroId, sTipoErro)
        
asReg(iIndice) = "6 - Rotina_Erro_Carregar_Msg. sErroId  = " & sErroId & " sTipoErro = " & sTipoErro
iIndice = iIndice + 1
        
        If Not IsMissing(vParam10) Then

asReg(iIndice) = "7 - Rotina_Erro_Carregar (Not IsMissing(vParam10))"
iIndice = iIndice + 1

            lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
            
asReg(iIndice) = "8 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam10)) = " & CStr(lErroAux)
iIndice = iIndice + 1

            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam9) Then
            
asReg(iIndice) = "9 - Rotina_Erro_Carregar (Not IsMissing(vParam9))"
iIndice = iIndice + 1

                lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
                
asReg(iIndice) = "10 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam9)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam8) Then
                
asReg(iIndice) = "11 - Rotina_Erro_Carregar (Not IsMissing(vParam8))"
iIndice = iIndice + 1

                    lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                    
asReg(iIndice) = "12 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam8)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam7) Then
                    
asReg(iIndice) = "13 - Rotina_Erro_Carregar (Not IsMissing(vParam7))"
iIndice = iIndice + 1

                        lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                        
asReg(iIndice) = "14 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam7)) = " & CStr(lErroAux)
iIndice = iIndice + 1
                        
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam6) Then
                        
asReg(iIndice) = "15 - Rotina_Erro_Carregar (Not IsMissing(vParam6))"
iIndice = iIndice + 1

                            lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                            
asReg(iIndice) = "16 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam6)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam5) Then
                            
asReg(iIndice) = "17 - Rotina_Erro_Carregar (Not IsMissing(vParam5))"
iIndice = iIndice + 1

                                lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                                
asReg(iIndice) = "18 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam5)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam4) Then
                                
asReg(iIndice) = "19 - Rotina_Erro_Carregar (Not IsMissing(vParam4))"
iIndice = iIndice + 1

                                    lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                    
asReg(iIndice) = "20 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam4)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam3) Then

asReg(iIndice) = "21 - Rotina_Erro_Carregar (Not IsMissing(vParam3))"
iIndice = iIndice + 1

                                        lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                        
asReg(iIndice) = "22 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam3)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam2) Then
                                        
asReg(iIndice) = "23 - Rotina_Erro_Carregar (Not IsMissing(vParam2))"
iIndice = iIndice + 1

                                            lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                            
asReg(iIndice) = "24 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam2)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                            If lErroAux Then Exit Function
                                        Else
                                            If Not IsMissing(vParam1) Then
                                            
asReg(iIndice) = "25 - Rotina_Erro_Carregar (Not IsMissing(vParam1))"
iIndice = iIndice + 1

                                                lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                                
asReg(iIndice) = "26 - lErroAux (Rotina_Erro_Carregar Not IsMissing(vParam1)) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                                If lErroAux Then Exit Function
                                            Else
                                            
asReg(iIndice) = "27 - Rotina_Erro_Carregar Else)"
iIndice = iIndice + 1

                                                lErroAux = Rotina_Erro_Carregar(sTipoErro, 0, sErro, "", "", "", "", "", "", "", "", "", "")
                                                
asReg(iIndice) = "28 - lErroAux (Rotina_Erro_Carregar Else) = " & CStr(lErroAux)
iIndice = iIndice + 1

                                                If lErroAux Then Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
asReg(iIndice) = "29 - Set X = CreateObject('TelasAdm.ClassTelasAdm')"
iIndice = iIndice + 1
        
        Set X = CreateObject("TelasAdm.ClassTelasAdm")
        
asReg(iIndice) = "30 - X Is Nothing"
iIndice = iIndice + 1

        If X Is Nothing Then Error 58450

asReg(iIndice) = "31 - X.objFormMsgAviso Is Nothing"
iIndice = iIndice + 1
        
        If X.objFormMsgAviso Is Nothing Then Error 58451

asReg(iIndice) = "32 - X.objFormMsgAviso.sErro = sErro = " & CStr(sErro)
iIndice = iIndice + 1

        'Passa o Erro e o Tipo dos Botões
        X.objFormMsgAviso.sErro = sErro
        
asReg(iIndice) = "33 - X.objFormMsgAviso.MsgBoxTipo = MsgBoxTipo"
iIndice = iIndice + 1

        X.objFormMsgAviso.MsgBoxTipo = MsgBoxTipo
        
asReg(iIndice) = "34 - X.objFormMsgAviso.Show vbModal"
iIndice = iIndice + 1

        X.objFormMsgAviso.Show vbModal
        
asReg(iIndice) = "35 - Rotina_Aviso = X.objFormMsgAviso.MsgBoxResultado"
iIndice = iIndice + 1

        'Retorna o Resultado da escolha
        Rotina_Aviso = X.objFormMsgAviso.MsgBoxResultado
        
asReg(iIndice) = "36 - Rotina_Aviso = X.objFormMsgAviso.MsgBoxResultado = " & CStr(Rotina_Aviso)
iIndice = iIndice + 1

    '    Rotina_Aviso = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")
    End If
    
    Exit Function
    
Erro_Rotina_Aviso:

asReg(iIndice) = "37 - Err = " & CStr(Err) & " Error = " & Error
iIndice = iIndice + 1

asReg(iIndice) = "****************************************************************************"
iIndice = iIndice + 1

   
    Select Case Err
    
        Case 58450, 58451
            Call MsgBox("Erro na exibição de uma pergunta (ou aviso).", vbOKOnly, "SGE - Forprint")
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Call Grava_Log(asReg)
        
    Error 59060 'para que quem chamou nao prossiga "normalmente" na operacao
    
    Exit Function
    
End Function

Function Grid_Inicializa(objGridInt As AdmGrid)

Dim objCampo As Object
Dim iIndice As Integer
Dim vName As Variant
Dim lLargura As Long
Dim iAltura As Integer
Dim iCol As Integer
Dim iLin As Integer
Dim iFator As Integer

On Error GoTo Erro_Grid_Inicializa

    objGridInt.objGrid.HighLight = flexHighlightNever
    
    'desabilita temporariamente o redesenho
    objGridInt.objGrid.Redraw = False

    objGridInt.iTeclado = 0
    objGridInt.lErroSaidaCelula = 0
    objGridInt.iSaidaCelula = 0

    objGridInt.objGrid.Cols = objGridInt.colColuna.Count

    iCol = 0

    objGridInt.objGrid.ColAlignment(0) = 4
    
    For iIndice = 1 To objGridInt.objGrid.Cols - 1
        objGridInt.objGrid.ColAlignment(iIndice) = 0
    Next

    'novo
    'indica que rotinas como RowColChange não serão executadas
    'permite que se caminhe por todas as linhas do grid
    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO

    For Each vName In objGridInt.colCampo
        iCol = iCol + 1
        Set objCampo = objGridInt.objForm.Controls(vName)
        
        If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iCol))
        
        objCampo.left = POSICAO_FORA_TELA
        objCampo.TabStop = False
        'novo
        'se uma das colunas é uma checkbox inicializa todas as celulas
        'da coluna com a imagem de uma checkbox desmarcada
        
        If Controle_ObterNomeClasse(objCampo) = "CheckBox" Then
            objCampo.Visible = False
            objGridInt.objGrid.Col = iCol
            For iLin = 1 To objGridInt.objGrid.Rows - 1
                objGridInt.objGrid.Row = iLin
                objGridInt.objGrid.CellForeColor = objGridInt.objGrid.BackColor
                objGridInt.objGrid.ForeColorSel = objGridInt.objGrid.BackColor
                objGridInt.objGrid.Text = "0"
                objGridInt.objGrid.CellPicture = objGridInt.objCheckboxUnchecked
                
                If vName = "CTBGerencial" And objGridInt.iGerencial = 2 Then
                    objGridInt.objGrid.Text = objGridInt.iGerencial
                    objGridInt.objGrid.CellPicture = objGridInt.objCheckboxGrayed
                End If
                
                objGridInt.objGrid.CellPictureAlignment = flexAlignCenterCenter
            Next
        End If
        
        If Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
            objCampo.Visible = False
            objGridInt.objGrid.Col = iCol
            For iLin = 1 To objGridInt.objGrid.Rows - 1
                objGridInt.objGrid.Row = iLin
                objGridInt.objGrid.CellForeColor = objGridInt.objGrid.BackColor
                objGridInt.objGrid.ForeColorSel = objGridInt.objGrid.BackColor
                objGridInt.objGrid.Text = "0"
                objGridInt.objGrid.CellPicture = objGridInt.objOptionButtonUnchecked
                objGridInt.objGrid.CellPictureAlignment = flexAlignCenterCenter
            Next
        End If
        
        If Controle_ObterNomeClasse(objCampo) = "MaskEdBox" Then
            If InStr(objCampo.Format, "0.00") <> 0 Or objCampo.Format = FORMATO_ESTOQUE Or objCampo.Format = FORMATO_CUSTO Or objCampo.Format = "0%" Then objGridInt.objGrid.ColAlignment(iCol) = flexAlignRightTop
        End If
        
        
    Next
    
    'retorna a linha e coluna atuais
'    If objGridInt.objGrid.Row <> iLinhaAtual Then objGridInt.objGrid.Row = iLinhaAtual
'    If objGridInt.objGrid.Col <> iColAtual Then objGridInt.objGrid.Col = iColAtual
    
    objGridInt.objGrid.Row = 0
    objGridInt.objGrid.Col = 0
    
    If objGridInt.iIncluirBotao = GRID_INCLUIR_BOTAO Then
        objGridInt.objGrid.CellPicture = objGridInt.objButton
        objGridInt.objGrid.CellPictureAlignment = flexAlignCenterCenter
    End If
    
    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
    
    For iIndice = 0 To objGridInt.objGrid.Cols - 1
        objGridInt.objGrid.FixedAlignment(iIndice) = flexAlignCenterCenter
        objGridInt.objGrid.TextMatrix(0, iIndice) = objGridInt.colColuna.Item(iIndice + 1)
    Next
    
    iAltura = 0
    For iIndice = 1 To objGridInt.colCampo.Count
        Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(iIndice))
        
        If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iIndice))
            
        objGridInt.objGrid.ColWidth(iIndice) = objCampo.Width + 15
        If objCampo.Height > iAltura Then iAltura = objCampo.Height
    Next
    
    iAltura = iAltura + 10
    
    For iIndice = 0 To objGridInt.objGrid.Rows - 1
        objGridInt.objGrid.RowHeight(iIndice) = iAltura
    Next
    
    iAltura = 100
    For iIndice = 0 To objGridInt.iLinhasVisiveis + objGridInt.objGrid.FixedRows - 1
        iAltura = iAltura + objGridInt.objGrid.RowHeight(iIndice) + objGridInt.objGrid.GridLineWidth
    Next
    
    
    objGridInt.objGrid.Height = iAltura
    
    If objGridInt.objGrid.Rows <= objGridInt.iLinhasVisiveis + objGridInt.objGrid.FixedRows Then
        lLargura = 0
    Else
        lLargura = 300
    End If
    
    For iIndice = 0 To objGridInt.objGrid.Cols - 1
        lLargura = lLargura + objGridInt.objGrid.ColWidth(iIndice) + 15 + objGridInt.objGrid.GridLineWidth
    Next
    
    If objGridInt.iGridLargAuto = 1 Then
        objGridInt.objGrid.Width = lLargura + 40
    End If
    
    If lLargura > objGridInt.objGrid.Width Then
        objGridInt.objGrid.Height = objGridInt.objGrid.Height + 240
    End If
    
    For iIndice = 1 To objGridInt.objGrid.Rows - 1
        objGridInt.objGrid.TextMatrix(iIndice, 0) = CStr(iIndice)
    Next
    
    Call CF("TelaControlLargUsu_Le_Fator", objGridInt.objForm.Name, iFator)
    If iFator > 0 Then Call CF("Trata_Tamanho_Grid", objGridInt, 0, iFator)
    
    'habilita o redesenho
    objGridInt.objGrid.Redraw = True
    
    Grid_Inicializa = SUCESSO
    
    Exit Function
    
Erro_Grid_Inicializa:
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142008)

    End Select

    Exit Function
        
End Function

'Public Function Grid_Reseta2(objGridInt As AdmGrid)
'
'Dim objCampo As Control
'Dim vName As Variant
'Dim lLargura As Long
'Dim iAltura As Integer
'Dim iCol As Integer
'Dim iLin As Integer
'
'On Error GoTo Erro_Grid_Reseta
'
'    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO
'    objGridInt.objGrid.Redraw = False
'
'    For iCol = 1 To objGridInt.objGrid.Cols - 1
'
'        For iLin = objGridInt.objGrid.FixedRows To objGrid.Rows - 1
'
'            objGridInt.objGrid.Col = iCol
'            objGridInt.objGrid.Row = iLin
'            Set objGridInt.objGrid.CellPicture = Nothing
'            objGridInt.objGrid.Text = ""
'
'
'        Next
'
'    Next
'
'    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
'    objGridInt.objGrid.Redraw = True
'
'    Grid_Reseta2 = SUCESSO
'
'    Exit Function
'
'Erro_Grid_Reseta2:
'
'    Grid_Reseta2 = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142009)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Grid_Reseta(objGridInt As AdmGrid) As Long

Dim objCampo As Object
Dim vName As Variant
Dim lLargura As Long
Dim iAltura As Integer
Dim iCol As Integer
Dim iLin As Integer, iRows As Integer
Dim vForeColorDefault As Variant, vForeColorSelDefault As Variant, vCellPictureAlignmentDefault As Variant, vColAlignment As Variant

On Error GoTo Erro_Grid_Reseta

    'desabilita temporariamente o redesenho
    objGridInt.objGrid.Redraw = False
    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO

    'crio uma linha-coluna ficticia para obter valores default
    objGridInt.objGrid.Rows = objGridInt.objGrid.Rows + 1
    objGridInt.objGrid.Cols = objGridInt.objGrid.Cols + 1
    objGridInt.objGrid.Row = objGridInt.objGrid.Rows - 1
    objGridInt.objGrid.Col = objGridInt.objGrid.Cols - 1
    vForeColorDefault = CLng(objGridInt.objGrid.CellForeColor)
    vForeColorSelDefault = CLng(objGridInt.objGrid.ForeColorSel)
    vCellPictureAlignmentDefault = CLng(objGridInt.objGrid.CellPictureAlignment)
    vColAlignment = CLng(objGridInt.objGrid.ColAlignment(objGridInt.objGrid.Col))
    objGridInt.objGrid.Rows = objGridInt.objGrid.Rows - 1
    objGridInt.objGrid.Cols = objGridInt.objGrid.Cols - 1
    
    For iCol = 1 To objGridInt.objGrid.Cols - 1

        objGridInt.objGrid.Col = iCol

        For iLin = objGridInt.objGrid.FixedRows To objGridInt.objGrid.Rows - 1
        
            objGridInt.objGrid.Row = iLin
            
            'se estes atributos nao foram alterados entao a coluna nao era checkbox, nem option button...
            If CLng(objGridInt.objGrid.CellForeColor) = vForeColorDefault And CLng(objGridInt.objGrid.ForeColorSel) = vForeColorSelDefault And CLng(objGridInt.objGrid.CellPictureAlignment) = vCellPictureAlignmentDefault Then Exit For
                
            objGridInt.objGrid.CellForeColor = vForeColorDefault
            objGridInt.objGrid.ForeColorSel = vForeColorSelDefault
            objGridInt.objGrid.CellPictureAlignment = vCellPictureAlignmentDefault
            If Not (objGridInt.objGrid.CellPicture Is Nothing) Then Set objGridInt.objGrid.CellPicture = Nothing
            If objGridInt.objGrid.Text <> "" Then objGridInt.objGrid.Text = ""
            
        Next
    
        objGridInt.objGrid.ColAlignment(iCol) = vColAlignment
        
    Next
        
    objGridInt.objGrid.Redraw = True
    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
    
    Grid_Reseta = SUCESSO
    
    Exit Function
    
Erro_Grid_Reseta:
    
    Grid_Reseta = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142010)

    End Select

    Exit Function
        
End Function

Private Function Grid_Campo_Libera_Foco1(objGridInt As AdmGrid, iCriticaCampo As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Grid_Campo_Libera_Foco1


    If objGridInt.lErroSaidaCelula <> 0 Then Error 5691

    If objGridInt.iSaidaCelula = 1 And objGridInt.iTeclado = 0 Then
    
        iCriticaCampo = 1
        
    Else

        iCriticaCampo = 0
        
    End If
    
    Grid_Campo_Libera_Foco1 = SUCESSO
    
    Exit Function
    
Erro_Grid_Campo_Libera_Foco1:

    Grid_Campo_Libera_Foco1 = Err

    Select Case Err

        Case 5691
            Call Grid_Retorna_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142011)

    End Select

    Exit Function

End Function

Private Function Grid_Campo_Libera_Foco2(objGridInt As AdmGrid) As Long

    objGridInt.objGrid.FocusRect = flexFocusHeavy
    
    objGridInt.iSaidaCelula = 0
    objGridInt.iEscEnter = 0
    
    Grid_Campo_Libera_Foco2 = SUCESSO

End Function

Function Grid_Abandona_Celula(objGridInt As AdmGrid) As Long

Dim dPercentual As Double

    If objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then

        If Not (Controle_ObterNomeClasse(objGridInt.objControle) = "CheckBox" Or Controle_ObterNomeClasse(objGridInt.objControle) = "OptionButton") Then
    
            If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
                Call objGridInt.objForm.ROTINA_GRID_ENABLE(objGridInt.objGrid.Row, objGridInt.objControle, ROTINA_GRID_ABANDONA_CELULA)
            End If
    
            If objGridInt.objControle.Enabled = True Then
        
                objGridInt.objControle.TabStop = False
                objGridInt.objControle.left = POSICAO_FORA_TELA
                
                'Se for uma masked edit com mascara de data
                If Controle_ObterNomeClasse(objGridInt.objControle) = "MaskEdBox" Then
                    If objGridInt.objControle.Mask = "##/##/##" Then
                        If objGridInt.objControle.Text = "  /  /  " Then
                            objGridInt.objGrid.Text = ""
                        Else
                            objGridInt.objGrid.Text = Format(objGridInt.objControle.Text, "dd/mm/yyyy")
                        End If
                    ElseIf objGridInt.objControle.Format = "#,##0.00" Then
                        If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then
                            objGridInt.objGrid.Text = Format(objGridInt.objControle.Text, "Standard")
                        Else
                            objGridInt.objGrid.Text = ""
                        End If
                    ElseIf objGridInt.objControle.Format = "0%" Then
                        If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then
                            dPercentual = CDbl(objGridInt.objControle.Text) / 100
'                            objGridInt.objGrid.Text = Format(objGridInt.objControle.Text, "##0%")
                            objGridInt.objGrid.Text = Format(dPercentual, "Percent")
'                            objGridInt.objGrid.Text = Format(objGridInt.objControle.Text, "##0\%")
                        Else
                            objGridInt.objGrid.Text = ""
                        End If
                    ElseIf objGridInt.objControle.Format = FORMATO_ESTOQUE Then
                        If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then
                            objGridInt.objGrid.Text = Formata_Estoque(CDbl(objGridInt.objControle.Text))
                        Else
                            objGridInt.objGrid.Text = ""
                        End If
                    ElseIf objGridInt.objControle.Format = FORMATO_CUSTO Then
                        If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then
                            objGridInt.objGrid.Text = Formata_Custo(CDbl(objGridInt.objControle.Text))
                        Else
                            objGridInt.objGrid.Text = ""
                        End If
                    Else
                        If Len(objGridInt.objControle.ClipText) > 0 Then
                            objGridInt.objGrid.Text = IIf(objGridInt.objControle.Format <> "", Format(objGridInt.objControle.Text, objGridInt.objControle.Format), objGridInt.objControle.Text)
                        Else
                            objGridInt.objGrid.Text = ""
                        End If
                    End If
                Else
                    'se for qualquer outro campo somente copia o dado
                    objGridInt.objGrid.Text = objGridInt.objControle.Text
                End If
            End If
        End If
    End If
    
    Grid_Abandona_Celula = SUCESSO
    
End Function

Function Grid_Finaliza_Saida_Celula(objGridInt As AdmGrid) As Long

    objGridInt.iSaidaCelula = 0
    objGridInt.iExecutaSaidaCelula = 0
    Grid_Finaliza_Saida_Celula = SUCESSO
    
End Function

Function Grid_Inicializa_Saida_Celula(objGridInt As AdmGrid) As Long

Dim objCampo As Object, objGrid As Object
   
    Set objGrid = objGridInt.objGrid
   
    With objGridInt
        If .lErroSaidaCelula = 0 And objGrid.Row > 0 And objGrid.Col > 0 And .iTeclado = 0 And objGrid.Row - objGrid.FixedRows <= .iLinhasExistentes - .iProibidoIncluir And .iExecutaSaidaCelula = 1 And .iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
            
            Set objCampo = .objForm.Controls(.colCampo(.objGrid.Col))
            
            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(.colIndex(.objGrid.Col))
            
            If objCampo.Enabled = True Then
        
                Grid_Inicializa_Saida_Celula = SUCESSO
                'indica que começou a fazer o tratamento de saida de celula
                objGridInt.iSaidaCelula = 4
            Else
                Grid_Inicializa_Saida_Celula = 5690
            End If
        Else
            Grid_Inicializa_Saida_Celula = 5690
        End If
    End With
End Function

Function Grid_Trata_Erro_Saida_Celula(objGridInt As AdmGrid)

    objGridInt.iLinhaAntiga = objGridInt.objGrid.Row
    objGridInt.iColunaAntiga = objGridInt.objGrid.Col
    objGridInt.lErroSaidaCelula = 1
    objGridInt.iSaidaCelula = 1
    Grid_Trata_Erro_Saida_Celula = SUCESSO
            
End Function


Function Grid_Trata_Erro_Saida_Celula1(objGridInt As AdmGrid, ByVal iLinha As Integer, ByVal iColuna As Integer)
'Ao constatar o erro coloca o foco em outro controle dentro de outra celula

    Call Grid_Abandona_Celula(objGridInt)
    Set objGridInt.objControle = objGridInt.objForm.Controls(objGridInt.colCampo(iColuna))
    
    objGridInt.iLinhaAntiga = iLinha
    objGridInt.iColunaAntiga = iColuna
    objGridInt.lErroSaidaCelula = 1
    objGridInt.iSaidaCelula = 0
    Grid_Trata_Erro_Saida_Celula1 = SUCESSO
            
End Function

Function Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt As AdmGrid)

Dim lErro As Long

    objGridInt.iLinhaAntiga = objGridInt.objGrid.Row
    objGridInt.iColunaAntiga = objGridInt.objGrid.Col
    'lErroSaidaCelula era igual a 2. Com o MDI passou para 3
    objGridInt.lErroSaidaCelula = 3
    If objGridInt.objForm.ActiveControl Is objGridInt.objControle And objGridInt.iEscEnter = 0 Then objGridInt.objGrid.SetFocus
    'retirado devido ao MDI em 24/9/99. Tenta evitar que quando a tela chamada já está visivel não apareça a mensagem de erro 2 vezes. Vai habilitar nas rotinas que tratam o erro da Saida de Celula.
    'If objGridInt.objControle.Visible = True Then objGridInt.objControle.SetFocus

    Grid_Trata_Erro_Saida_Celula_Chama_Tela = SUCESSO
            
End Function

Function Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt As AdmGrid)

Dim lErro As Long

    objGridInt.iLinhaAntiga = objGridInt.objGrid.Row
    objGridInt.iColunaAntiga = objGridInt.objGrid.Col
    objGridInt.lErroSaidaCelula = 3
    If objGridInt.objForm.ActiveControl Is objGridInt.objControle And objGridInt.iEscEnter = 0 Then objGridInt.objGrid.SetFocus
    'retirado devido ao MDI em 24/9/99. Tenta evitar que quando a tela chamada já está visivel não apareça a mensagem de erro 2 vezes. Vai habilitar nas rotinas que tratam o erro da Saida de Celula.
    'If objGridInt.objControle.Visible = True Then objGridInt.objControle.SetFocus

    Grid_Trata_Erro_Saida_Celula_Chama_Tela1 = SUCESSO
            
End Function

Function Grid_RowColChange(objGridInt As AdmGrid) As Long

    If objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
        'trata o posicionamento do grid se detectou um erro em uma celula
        If objGridInt.lErroSaidaCelula <> 0 Then
    
            If objGridInt.objGrid.Row <> objGridInt.iLinhaAntiga Then
                objGridInt.objGrid.Row = objGridInt.iLinhaAntiga
            ElseIf objGridInt.objGrid.Col <> objGridInt.iColunaAntiga Then
                objGridInt.objGrid.Col = objGridInt.iColunaAntiga
            Else
                If objGridInt.objGrid.Row < objGridInt.objGrid.TopRow Then
                    objGridInt.objGrid.TopRow = objGridInt.objGrid.Row
                ElseIf objGridInt.objGrid.Row > objGridInt.objGrid.TopRow + 4 Then
                    objGridInt.objGrid.TopRow = objGridInt.objGrid.Row - 4
                End If
                If objGridInt.lErroSaidaCelula <> 0 Then
                
                    If objGridInt.lErroSaidaCelula = 1 And objGridInt.objControle.Visible = True Then objGridInt.objControle.SetFocus
                    
                    objGridInt.objControle.left = objGridInt.objGrid.left + objGridInt.objGrid.CellLeft
                    objGridInt.objControle.top = objGridInt.objGrid.top + objGridInt.objGrid.CellTop
                
                End If
                If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then
                    objGridInt.lErroSaidaCelula = 0
                End If
            End If
            
            'novo
        Else
        
            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows > objGridInt.iLinhasExistentes - objGridInt.iProibidoIncluir Then
                objGridInt.objGrid.Row = objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - objGridInt.iProibidoIncluir
            End If
        End If
    End If
    
    Grid_RowColChange = SUCESSO

End Function

Function Grid_Retorna_Celula(objGridInt As AdmGrid) As Long
        
        'mario
        If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then
        
            If objGridInt.objGrid.Row <> objGridInt.iLinhaAntiga Then
                objGridInt.objGrid.Row = objGridInt.iLinhaAntiga
            End If
            
            If objGridInt.objGrid.Col <> objGridInt.iColunaAntiga Then
                objGridInt.objGrid.Col = objGridInt.iColunaAntiga
            End If
            
            If objGridInt.objGrid.Row < objGridInt.objGrid.TopRow Then
                objGridInt.objGrid.TopRow = objGridInt.objGrid.Row
            ElseIf objGridInt.objGrid.Row > objGridInt.objGrid.TopRow + 4 Then
                objGridInt.objGrid.TopRow = objGridInt.objGrid.Row - 4
            End If
            
            If objGridInt.lErroSaidaCelula = 1 And objGridInt.objControle.Visible = True Then objGridInt.objControle.SetFocus
            objGridInt.lErroSaidaCelula = 0
        
        
        End If
        
        'colocado por mario em 24/9/99 devido ao MDI
        If objGridInt.lErroSaidaCelula = 3 Then
            objGridInt.lErroSaidaCelula = 0
        End If
        
        
        Grid_Retorna_Celula = SUCESSO

End Function

Function Grid_Campo_Recebe_Foco(objGridInt As AdmGrid) As Long

    If objGridInt.lErroSaidaCelula = 3 Then
        objGridInt.lErroSaidaCelula = 0
    End If

    If objGridInt.lErroSaidaCelula = 2 Then objGridInt.lErroSaidaCelula = 3

    If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then
        objGridInt.iSaidaCelula = 1
    End If
    
    'mario foco
    objGridInt.iTeclado = 0
    
    Grid_Campo_Recebe_Foco = SUCESSO
End Function

Function Grid_Campo_Recebe_Foco_Contabil(objGridInt As AdmGrid) As Long

    If objGridInt.lErroSaidaCelula = 3 Then
        objGridInt.lErroSaidaCelula = 0
    End If

    If objGridInt.lErroSaidaCelula = 2 Then objGridInt.lErroSaidaCelula = 3


'    If objGridInt.lErroSaidaCelula = 2 Or objGridInt.lErroSaidaCelula = 3 Then objGridInt.lErroSaidaCelula = 0

    If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then
        objGridInt.iSaidaCelula = 1
    End If
    
    'mario foco
    objGridInt.iTeclado = 0
    
    Grid_Campo_Recebe_Foco_Contabil = SUCESSO
    
End Function

Function Grid_Recebe_Foco(objGridInt As AdmGrid) As Long

'    If objGridInt.lErroSaidaCelula = 2 Then
'        objGridInt.objControle.SetFocus
'        objGridInt.lErroSaidaCelula = 0
'    End If

'mario foco
    If objGridInt.iEscEnter = 1 Then
        objGridInt.iEscEnter = 0
    Else
        objGridInt.iTeclado = 1
    End If
    
    'mario
    If objGridInt.objGrid.Row = 0 And objGridInt.objGrid.Col = 0 Then
        objGridInt.objGrid.Row = 1
        objGridInt.objGrid.Col = 1
        objGridInt.objGrid.FocusRect = flexFocusHeavy
    End If
    
    Grid_Recebe_Foco = SUCESSO
End Function

Function Grid_Libera_Foco(objGridInt As AdmGrid) As Long
    If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then objGridInt.iTeclado = 0
    Grid_Libera_Foco = SUCESSO
End Function

Function Grid_Libera_Foco_Contabil(objGridInt As AdmGrid) As Long
    If objGridInt.lErroSaidaCelula <> 2 And objGridInt.lErroSaidaCelula <> 3 Then objGridInt.iTeclado = 0
    Grid_Libera_Foco_Contabil = SUCESSO
End Function

Function Grid_Click(objGridInt As AdmGrid, iExecutaEntradaCelula As Integer) As Long

Dim objCampo As Object
Dim iLin As Integer
Dim iLinCorrente As Integer, iMouseRow As Integer

    If objGridInt.objForm.Enabled = True Then

        iExecutaEntradaCelula = 0
        
        If objGridInt.objGrid.Col > 0 Then
        
            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(objGridInt.objGrid.Col))
            
            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(objGridInt.objGrid.Col))
            
            If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
                Call objGridInt.objForm.ROTINA_GRID_ENABLE(objGridInt.objGrid.Row, objCampo, ROTINA_GRID_CLICK)
            End If
            
            If objCampo.Enabled = True Then
            
                iMouseRow = objGridInt.objGrid.MouseRow
                If iMouseRow <> 0 Then iMouseRow = objGridInt.objGrid.Row
                
                'novo
                'se o campo é um checkbox
                If Controle_ObterNomeClasse(objCampo) = "CheckBox" Then
                    
                    If iMouseRow > 0 And objGridInt.objGrid.Col > 0 And iMouseRow - objGridInt.objGrid.FixedRows <= objGridInt.iLinhasExistentes - objGridInt.iProibidoIncluir Then
                        'se o campo está desligado ==> ligar
                        If objGridInt.objGrid.Text = "0" Or objGridInt.objGrid.Text = "" Then
                            objGridInt.objGrid.Text = "1"
                            objGridInt.objGrid.CellPicture = objGridInt.objCheckboxChecked
                        Else
                            If giContabGerencial = 1 And objCampo.Tag = "1" Then
                                If objGridInt.objGrid.Text = "1" Then
                                    objGridInt.objGrid.Text = "2"
                                    objGridInt.objGrid.CellPicture = objGridInt.objCheckboxGrayed
                                ElseIf objGridInt.objGrid.Text = "2" Then
                                    'se o campo está ligado ==> desligar
                                    objGridInt.objGrid.Text = "0"
                                    objGridInt.objGrid.CellPicture = objGridInt.objCheckboxUnchecked
                                End If
                            Else
                                'se o campo está ligado ==> desligar
                                objGridInt.objGrid.Text = "0"
                                objGridInt.objGrid.CellPicture = objGridInt.objCheckboxUnchecked
                            End If
                        End If
                        If objCampo.Value <> 0 Then
                            objCampo.Value = 0
                        Else
                            objCampo.Value = 1
                        End If
                        
                        objGridInt.objGrid.FocusRect = flexFocusHeavy
                    End If
                        
                'se o campo é um optionbutton
                ElseIf Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
                
                    If iMouseRow > 0 And objGridInt.objGrid.Col > 0 And iMouseRow - objGridInt.objGrid.FixedRows <= objGridInt.iLinhasExistentes - objGridInt.iProibidoIncluir Then
                        'se o campo está desligado ==> ligar
                        If objGridInt.objGrid.Text = "0" Or objGridInt.objGrid.Text = "" Then
                            
                            'desabilita temporariamente o redesenho
                            objGridInt.objGrid.Redraw = False
                        
                            iLinCorrente = objGridInt.objGrid.Row
                            For iLin = 1 To objGridInt.objGrid.Rows - 1
                                If objGridInt.objGrid.TextMatrix(iLin, objGridInt.objGrid.Col) = "1" Then
                                    objGridInt.objGrid.Row = iLin
                                    objGridInt.objGrid.Text = "0"
                                    objGridInt.objGrid.CellPicture = objGridInt.objOptionButtonUnchecked
                                End If
                            Next
                            
                            If iLinCorrente <> objGridInt.objGrid.Row Then objGridInt.objGrid.Row = iLinCorrente
                            objGridInt.objGrid.Text = "1"
                            objGridInt.objGrid.CellPicture = objGridInt.objOptionButtonChecked
                            
                            If objCampo.Value <> False Then
                                objCampo.Value = False
                            End If
                            
                            objCampo.Value = True
                            
                            'reabilita o redesenho
                            objGridInt.objGrid.Redraw = True
                            
                        End If
                        
                        objGridInt.objGrid.FocusRect = flexFocusHeavy
                    End If
                        
                Else
            
                    objGridInt.objGrid.FocusRect = flexFocusNone
            
                    If objGridInt.iTeclado = 1 Then
                        objGridInt.iTeclado = 0
                        iExecutaEntradaCelula = 1
                    End If
                End If
            Else
                objGridInt.objGrid.FocusRect = flexFocusHeavy
                If objGridInt.objGrid.Visible = True Then objGridInt.objGrid.SetFocus
            End If
            
        End If
        
        'se for a coluna 0
        'verifica se o botao está ligado
        'Se estiver chama a funcao da tela GridBotao_Click
        If objGridInt.objGrid.MouseCol < objGridInt.objGrid.FixedCols And objGridInt.objGrid.MouseRow >= objGridInt.objGrid.FixedRows And objGridInt.iIncluirBotao = GRID_INCLUIR_BOTAO Then Call objGridInt.objForm.GridBotao_Click(objGridInt)
    
    End If
    Grid_Click = SUCESSO

End Function

Function Grid_Trata_Tecla(KeyAscii As Integer, objGridInt As AdmGrid, iExecutaEntradaCelula As Integer)

Dim objCampo As Object
Dim iLin As Integer
Dim iLinCorrente As Integer

    iExecutaEntradaCelula = 0
    
    If objGridInt.objGrid.Col > 0 Then
    
        If KeyAscii = vbKeyReturn Then
        
            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(objGridInt.objGrid.Col))
            
            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(objGridInt.objGrid.Col))
            
            'novo
            'se ao teclar enter o campo em questão for uma checkbox ou um optionbutton
            'irá passar para o proximo campo habilitado
            If Controle_ObterNomeClasse(objCampo) = "CheckBox" Or Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
                Call Grid_Trata_Tecla_Campo2(vbKeyReturn, objGridInt)
                
            Else
                objGridInt.iTeclado = 0
                KeyAscii = 0
                iExecutaEntradaCelula = 1
            End If
            
            If objCampo.Enabled = False Then
            
                'Call Grava_Log2("FUNÇÃO: Grid_Trata_Tecla objCampo.Enabled = False")
                'Call Grava_Log2("objGridInt: col " & CStr(objGridInt.objGrid.Col) & "; cols " & CStr(objGridInt.objGrid.Cols))
                'Call Grava_Log2("objGridInt: row " & CStr(objGridInt.objGrid.Row) & "; rows " & CStr(objGridInt.objGrid.Rows))
                'Call Grava_Log2("objGridInt: iLinhasExistentes " & CStr(objGridInt.iLinhasExistentes) & "; iLinhaAntiga " & CStr(objGridInt.iLinhaAntiga))
                'Call Grava_Log2("objCampo: Name " & objCampo.Name)
                'Call Grava_Log2(" ")
            
                 Call Grid_Trata_Tecla_Campo2(vbKeyReturn, objGridInt)
            End If
            
            
        'novo
        'se teclou a barra de espaço
        ElseIf KeyAscii = vbKeySpace Then
            
            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(objGridInt.objGrid.Col))
            
            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(objGridInt.objGrid.Col))
            
            If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
                Call objGridInt.objForm.ROTINA_GRID_ENABLE(objGridInt.objGrid.Row, objCampo, ROTINA_GRID_TRATA_TECLA)
            End If
            
            'se o campo é um checkbox, está habilitado e a linha que está tratando não é uma linha fixa
            If Controle_ObterNomeClasse(objCampo) = "CheckBox" And objCampo.Enabled = True And objGridInt.objGrid.Row >= objGridInt.objGrid.FixedRows Then
            
    '            objGridInt.objGrid.CellForeColor = objGridInt.objGrid.BackColor
    '            objGridInt.objGrid.ForeColorSel = objGridInt.objGrid.BackColor
            
                'se o campo está desligado ==> ligar
                If objGridInt.objGrid.Text = "0" Or objGridInt.objGrid.Text = "" Then
                    objGridInt.objGrid.CellPicture = objGridInt.objCheckboxChecked
                    objGridInt.objGrid.Text = "1"
                Else
                    If giContabGerencial = 1 And objCampo.Tag = "1" Then
                        If objGridInt.objGrid.Text = "1" Then
                            objGridInt.objGrid.Text = "2"
                            objGridInt.objGrid.CellPicture = objGridInt.objCheckboxGrayed
                        ElseIf objGridInt.objGrid.Text = "2" Then
                            'se o campo está ligado ==> desligar
                            objGridInt.objGrid.Text = "0"
                            objGridInt.objGrid.CellPicture = objGridInt.objCheckboxUnchecked
                        End If
                    Else
                
                        'se o campo está ligado ==> desligar
                        objGridInt.objGrid.Text = "0"
                        objGridInt.objGrid.CellPicture = objGridInt.objCheckboxUnchecked
                    End If
                End If
            
                If objCampo.Value <> 0 Then
                    objCampo.Value = 0
                Else
                    objCampo.Value = 1
                End If
                
            End If
            
            'se o campo é um optionbutton
            If Controle_ObterNomeClasse(objCampo) = "OptionButton" And objCampo.Enabled = True Then
            
                'se o campo está desligado ==> ligar
                If objGridInt.objGrid.Text = "0" Or objGridInt.objGrid.Text = "" Then
                
                    'desabilita temporariamente o redesenho
                    objGridInt.objGrid.Redraw = False
                
                    iLinCorrente = objGridInt.objGrid.Row
                    For iLin = 1 To objGridInt.objGrid.Rows - 1
                        If objGridInt.objGrid.TextMatrix(iLin, objGridInt.objGrid.Col) = "1" Then
                            objGridInt.objGrid.Row = iLin
                            objGridInt.objGrid.Text = "0"
                            objGridInt.objGrid.CellPicture = objGridInt.objOptionButtonUnchecked
                        End If
                    Next
                        
                    If iLinCorrente <> objGridInt.objGrid.Row Then objGridInt.objGrid.Row = iLinCorrente
                    objGridInt.objGrid.Text = "1"
                    objGridInt.objGrid.CellPicture = objGridInt.objOptionButtonChecked
                        
                    If objCampo.Value <> False Then
                        objCampo.Value = False
                    End If
                    
                    objCampo.Value = True
                        
                    'reabilita o redesenho
                    objGridInt.objGrid.Redraw = True
                        
                End If
            End If
        End If
    End If
    
    Grid_Trata_Tecla = SUCESSO
    
End Function

Function Grid_Trata_Tecla1(KeyCode As Integer, objGridInt As AdmGrid) As Long

Dim vbMsgRet As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_Grid_Trata_Tecla1

    If KeyCode = vbKeyDelete And objGridInt.iProibidoExcluir = 0 Then
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows + 1 <= objGridInt.iLinhasExistentes Then
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_LINHA_GRID", objGridInt.objGrid.Row)
            If vbMsgRet = vbYes Then
                lErro = Grid_Exclui_Linha0(objGridInt)
            Else
                objGridInt.objGrid.SetFocus
                Error 20627
            End If
        Else
            Error 20626
        End If
    ElseIf KeyCode = vbKeyInsert And objGridInt.iProibidoIncluirNoMeioGrid = 0 Then
        Call Grid_Inclui_Linha(objGridInt)
    ElseIf KeyCode = KEYCODE_LOCALIZAGRID Then
        Call CF("LocalizarGrid_Abre", objGridInt)
    ElseIf KeyCode = KEYCODE_EXPORTAGRIDEXCEL Then
        Call CF("Excel_Gera_Planilha_Grid", objGridInt)
    ElseIf KeyCode = vbKeyAdd Then
        Call CF("Trata_Tamanho_Grid", objGridInt)
    ElseIf KeyCode = vbKeySubtract Then
        Call CF("Trata_Tamanho_Grid", objGridInt, 1)
    Else
        Error 20625
    End If
        
    Grid_Trata_Tecla1 = SUCESSO
    
    Exit Function
    
Erro_Grid_Trata_Tecla1:
    
    Grid_Trata_Tecla1 = Err
    
    Select Case Err
    
        Case 20625, 20626, 20627
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142012)
    
    End Select
    
    Exit Function
        
End Function

Function Grid_Trata_Tecla2(KeyCode As Integer, objGridInt As AdmGrid)

Dim vbMsgRet As VbMsgBoxResult
Dim lErro As Long

    If KeyCode = vbKeyDelete And objGridInt.iProibidoExcluir = 0 Then
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows + 1 <= objGridInt.iLinhasExistentes Then
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_LINHA_GRID", objGridInt.objGrid.Row)
            If vbMsgRet = vbYes Then
                lErro = Grid_Exclui_Linha1(objGridInt)
            End If
        End If
    ElseIf KeyCode = KEYCODE_LOCALIZAGRID Then
        Call CF("LocalizarGrid_Abre", objGridInt)
    ElseIf KeyCode = KEYCODE_EXPORTAGRIDEXCEL Then
        Call CF("Excel_Gera_Planilha_Grid", objGridInt)
    ElseIf KeyCode = vbKeyAdd Then
        Call CF("Trata_Tamanho_Grid", objGridInt)
    ElseIf KeyCode = vbKeySubtract Then
        Call CF("Trata_Tamanho_Grid", objGridInt, 1)
    End If
        
    Grid_Trata_Tecla2 = SUCESSO
    
End Function

Function Grid_Trata_Tecla3(KeyCode As Integer, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Grid_Trata_Tecla3

    If KeyCode = vbKeyDelete And objGridInt.iProibidoExcluir = 0 Then
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows + 1 <= objGridInt.iLinhasExistentes Then
            Call Grid_Exclui_Linha0(objGridInt)
        Else
            Error 20626
        End If
    ElseIf KeyCode = KEYCODE_LOCALIZAGRID Then
        Call CF("LocalizarGrid_Abre", objGridInt)
    ElseIf KeyCode = KEYCODE_EXPORTAGRIDEXCEL Then
        Call CF("Excel_Gera_Planilha_Grid", objGridInt)
    ElseIf KeyCode = vbKeyAdd Then
        Call CF("Trata_Tamanho_Grid", objGridInt)
    ElseIf KeyCode = vbKeySubtract Then
        Call CF("Trata_Tamanho_Grid", objGridInt, 1)
    Else
        Error 20625
    End If
        
    Grid_Trata_Tecla3 = SUCESSO
    
    Exit Function
    
Erro_Grid_Trata_Tecla3:
    
    Grid_Trata_Tecla3 = Err
    
    Select Case Err
    
        Case 20625, 20626
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142013)
    
    End Select
    
    Exit Function
        
End Function

Private Function Grid_Exclui_Linha0(objGridInt As AdmGrid)

Dim iLinha As Integer
Dim iColuna As Integer
Dim iLinha1 As Integer

    For iLinha = objGridInt.objGrid.Row + 1 To objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - 1
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
            objGridInt.objGrid.TextMatrix(iLinha - 1, iColuna) = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
        Next
    Next
 
    For iColuna = 1 To objGridInt.objGrid.Cols - 1
        objGridInt.objGrid.TextMatrix(iLinha - 1, iColuna) = ""
    Next
 
    objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes - 1
 
    'novo
    'redesenha os checkbox
    Call Grid_Refresh_Checkbox(objGridInt)
 
    
End Function

Private Function Grid_Inclui_Linha(objGridInt As AdmGrid)

Dim iLinha As Integer
Dim iColuna As Integer
Dim iLinha1 As Integer

    For iLinha = objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - 1 To objGridInt.objGrid.Row Step -1
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
            objGridInt.objGrid.TextMatrix(iLinha + 1, iColuna) = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
        Next
    Next
 
    For iColuna = 1 To objGridInt.objGrid.Cols - 1
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iColuna) = ""
    Next
 
    objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
 
    'novo
    'redesenha os checkbox
    Call Grid_Refresh_Checkbox(objGridInt)
    
End Function

Private Function Grid_Exclui_Linha1(objGridInt As AdmGrid)

Dim iLinha As Integer
Dim iColuna As Integer
Dim iLinha1 As Integer

    For iLinha = objGridInt.objGrid.Row + 1 To objGridInt.objGrid.Rows - 1
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
            objGridInt.objGrid.TextMatrix(iLinha - 1, iColuna) = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
        Next
    Next
 
    objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes - 1
 
    'novo
    'redesenha os checkbox
    Call Grid_Refresh_Checkbox(objGridInt)
 
    
End Function

Function Grid_Exclui_Linha(objGridInt As AdmGrid, iLinha2 As Integer)

Dim iLinha As Integer
Dim iColuna As Integer
Dim iLinha1 As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_Exclui_Linha

    If iLinha2 < objGridInt.objGrid.FixedRows Or iLinha2 > objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - 1 Then Error 20319

    For iLinha = iLinha2 + 1 To objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - 1
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
            objGridInt.objGrid.TextMatrix(iLinha - 1, iColuna) = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
        Next
    Next
 
    For iColuna = 1 To objGridInt.objGrid.Cols - 1
        objGridInt.objGrid.TextMatrix(iLinha - 1, iColuna) = ""
    Next
 
    objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes - 1
 
    'novo
    'redesenha os checkbox
    Call Grid_Refresh_Checkbox(objGridInt)
 
    Grid_Exclui_Linha = SUCESSO
    
    Exit Function
    
Erro_Grid_Exclui_Linha:

    Grid_Exclui_Linha = Err
    
    Select Case Err
    
        Case 20319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_LINHA_INEXISTENTE", Err, iLinha2, objGridInt.objGrid.FixedRows, objGridInt.iLinhasExistentes + objGridInt.objGrid.FixedRows - 1)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142014)

    End Select
    
    Exit Function
    
End Function

Function Grid_Entrada_Celula(objGridInt As AdmGrid, iAlterado As Integer) As Long

Dim iIndice As Integer
Dim objCampo As Object
Dim dPercentual As Double
Dim iAlterado1 As Integer
Dim sTexto As String
    
On Error GoTo Erro_Grid_Entrada_Celula

    iAlterado1 = iAlterado
    
    If objGridInt.lErroSaidaCelula = 0 And objGridInt.iTeclado = 0 And objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
        
        If objGridInt.objGrid.Row > 0 And objGridInt.objGrid.Col > 0 And objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows <= objGridInt.iLinhasExistentes - objGridInt.iProibidoIncluir Then
        
            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(objGridInt.objGrid.Col))
            
            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(objGridInt.objGrid.Col))
            
            If Not (Controle_ObterNomeClasse(objCampo) = "CheckBox" Or Controle_ObterNomeClasse(objCampo) = "OptionButton") Then
            
                If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
                    Call objGridInt.objForm.ROTINA_GRID_ENABLE(objGridInt.objGrid.Row, objCampo, ROTINA_GRID_ENTRADA_CELULA)
                End If
                
                If objCampo.Enabled = True Then
                    objCampo.left = objGridInt.objGrid.left + objGridInt.objGrid.CellLeft
                    objCampo.top = objGridInt.objGrid.top + objGridInt.objGrid.CellTop
                    If Controle_ObterNomeClasse(objCampo) = "ComboBox" Then
                        objCampo.ListIndex = -1
                        If objCampo.Style = vbComboDropdown Then objCampo.Text = ""
                        If Len(objGridInt.objGrid.Text) > 0 Then
                            For iIndice = 0 To objCampo.ListCount - 1
                                If objCampo.List(iIndice) = objGridInt.objGrid.Text Then
                                    objCampo.ListIndex = iIndice
                                    Exit For
                                End If
                            Next
                            'Se não encontrou o texto em algum dos elementos da combo, coloca no texto da combo o texto do grid
                            If iIndice > objCampo.ListCount - 1 Then
                                objCampo.Text = objGridInt.objGrid.Text
                            End If
                        End If
                    ElseIf Controle_ObterNomeClasse(objCampo) = "MaskEdBox" Then

                        If Len(objGridInt.objGrid.Text) = 0 Then
                            If objCampo.PromptInclude = True Then
                                objCampo.PromptInclude = False
                                objCampo.Text = ""
                                objCampo.PromptInclude = True
                            Else
                                objCampo.Text = ""
                            End If
                        Else
                        
                            If objCampo.Mask = "##/##/##" Then
                                objCampo.Text = Format(objGridInt.objGrid.Text, "dd/mm/yy")
                            ElseIf objCampo.Format = "#,##0.00" Then
                                objCampo.Text = Format(objGridInt.objGrid.Text, "General Number")
                            ElseIf objCampo.Format = "0%" Then
                                dPercentual = CDbl(Format(objGridInt.objGrid.Text, "General Number")) * 100
                                objCampo.Text = CStr(dPercentual)
                            ElseIf objCampo.Format = FORMATO_ESTOQUE Then
                                objCampo.Text = Format(objGridInt.objGrid.Text, "General Number")
                            ElseIf objCampo.Format = FORMATO_CPF Or objCampo.Format = FORMATO_CGC Then
                                If Len(objGridInt.objGrid.Text) = 14 Then
                                    objCampo.Mask = FORMATO_CPF
                                Else
                                    objCampo.Mask = FORMATO_CGC
                                End If
                            
                                objCampo.Text = objGridInt.objGrid.Text
                                sTexto = objCampo.Text
                                objCampo.Mask = ""
                                objCampo.Text = sTexto
                            
                            Else
                                objCampo.Text = objGridInt.objGrid.Text
                            End If
                        End If
                    Else
                        objCampo.Text = objGridInt.objGrid.Text
                    End If
                    objCampo.TabStop = True
                    If objCampo.Visible = True Then objCampo.SetFocus
                    objGridInt.iExecutaSaidaCelula = 1
                End If
                
            End If
            
        End If
    
    End If

    If iAlterado <> iAlterado1 Then iAlterado = iAlterado1

    Grid_Entrada_Celula = SUCESSO

Erro_Grid_Entrada_Celula:

    Resume Next 'colocado em 15/05/2013 para tratar erros como o de tentar seleciona texto que nao existe em combo
    
End Function

Function Grid_Limpa_Campo(objCampo As Object) As Long

Dim objMasked As Object

    If Controle_ObterNomeClasse(objCampo) = "ComboBox" Then
    
        objCampo.ListIndex = -1
        
    ElseIf Controle_ObterNomeClasse(objCampo) = "MaskEdBox" Then
    
        Set objMasked = objCampo
        With objMasked
        
            If .PromptInclude = True Then
                .PromptInclude = False
                .Text = ""
                .PromptInclude = True
            Else
                .Text = ""
            End If
        End With
        
    ElseIf (Controle_ObterNomeClasse(objCampo) = "CheckBox" Or Controle_ObterNomeClasse(objCampo) = "OptionButton") Then
        'não faz nada
    Else
        objCampo.Text = ""
    End If

    'colocado por mario em 28/9/99 devido ao MDI
    objCampo.TabStop = False
    objCampo.left = POSICAO_FORA_TELA
    
    Grid_Limpa_Campo = SUCESSO

End Function


Function Grid_Scroll(objGridInt As AdmGrid) As Long

Dim objCampo As Object

    If objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
        
        objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO

        If objGridInt.objGrid.TopRow + objGridInt.iLinhasVisiveis - objGridInt.objGrid.FixedRows > objGridInt.iLinhasExistentes + 1 - objGridInt.iProibidoIncluir + objGridInt.iIncluirHScroll Then
            If objGridInt.iLinhasExistentes + 1 - objGridInt.iProibidoIncluir > objGridInt.iLinhasVisiveis - objGridInt.iIncluirHScroll Then
                objGridInt.objGrid.TopRow = objGridInt.iLinhasExistentes - objGridInt.iLinhasVisiveis + 2 - objGridInt.iProibidoIncluir + objGridInt.iIncluirHScroll
            Else
    '            objGridInt.objGrid.TopRow = objGridInt.objGrid.FixedRows + objGridInt.iIncluirHScroll
                objGridInt.objGrid.TopRow = objGridInt.objGrid.FixedRows
            End If
        End If
        
        If objGridInt.iTeclado = 0 Then
        
            If objGridInt.objGrid.Col > 0 Then
                   
                Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(objGridInt.objGrid.Col))
                
                If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(objGridInt.objGrid.Col))
                
                'se o campo atual é o campo que contém o foco
                If objCampo Is objGridInt.objForm.ActiveControl Then
                
                    If Not (Controle_ObterNomeClasse(objCampo) = "CheckBox" Or Controle_ObterNomeClasse(objCampo) = "OptionButton") Then
                    
                        If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
                            Call objGridInt.objForm.ROTINA_GRID_ENABLE(objGridInt.objGrid.Row, objCampo, ROTINA_GRID_SCROLL)
                        End If
                    
                        If objCampo.Enabled = True Then
            
                            If objGridInt.objGrid.Row >= objGridInt.objGrid.TopRow And objGridInt.objGrid.Row < objGridInt.objGrid.TopRow + objGridInt.iLinhasVisiveis And objGridInt.objGrid.ColIsVisible(objGridInt.objGrid.Col) Then
                                objCampo.left = objGridInt.objGrid.left + objGridInt.objGrid.CellLeft
                                objCampo.top = objGridInt.objGrid.top + objGridInt.objGrid.CellTop
                                objCampo.TabStop = True
                            Else
                                objCampo.TabStop = False
                                objCampo.left = POSICAO_FORA_TELA
                            End If
                        End If
                    End If
                End If
            End If
            
        End If
        
        objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
        
    End If
    
    Grid_Scroll = SUCESSO
    
End Function

Function Grid_Limpa(objGridInt As AdmGrid) As Long

Dim iLinha As Integer
Dim iColuna As Integer, objGrid As Object
Dim iExecutaRotinaEnable As Integer
Dim objCampo As Object
''Dim aRowColVazio() As String
''Dim aLinhaFixa() As String
''Dim aColFixa() As String

    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO

    objGridInt.iLinhasExistentes = 0

    Set objGrid = objGridInt.objGrid
    
    With objGrid
    
''        ReDim aRowColVazio(.Rows, .Cols)
''        ReDim aLinhaFixa(.Cols)
''        ReDim aColFixa(.Rows)
''
''        For iLinha = .FixedRows To .Rows - 1
''
''            aColFixa(iLinha) = .TextMatrix(iLinha, 0)
''
''        Next
''
''        For iColuna = .FixedCols To .Cols - 1
''
''            aLinhaFixa(iColuna) = .TextMatrix(0, iColuna)
''
''        Next
''
''        objGridInt.objGrid.TextMatrix = aRowColVazio
''
''        For iLinha = .FixedRows To .Rows - 1
''
''            .TextMatrix(iLinha, 0) = aColFixa(iLinha)
''
''        Next
''
''        For iColuna = .FixedCols To .Cols - 1
''
''            .TextMatrix(0, iColuna) = aLinhaFixa(iColuna)
''
''        Next
        
        .Row = .FixedRows
        .Col = .FixedCols
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .Text = ""
        .CellFontBold = False
        .FillStyle = flexFillSingle
        
''        For iLinha = .FixedRows To .Rows - 1
''
''            For iColuna = .FixedCols To .Cols - 1
''
''                .TextMatrix(iLinha, iColuna) = ""
''
''            Next
''
''        Next
    
    End With
        
    For iColuna = objGrid.FixedCols To objGrid.Cols - 1
    
        Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(iColuna))
    
        If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iColuna))
    
        Call Grid_Limpa_Campo(objCampo)
    Next
    
    'o codigo do rotina grid enable foi colocado para evitar a execução desta rotina quando da entrada ou
    'saida para a linha/coluna 0

    iExecutaRotinaEnable = objGridInt.iExecutaRotinaEnable

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    Call Grid_Refresh_Checkbox_Limpa(objGridInt)

    objGrid.Row = 0
    objGrid.Col = 0

    objGrid.TopRow = 1

    'retorna o valor anterior
    objGridInt.iExecutaRotinaEnable = iExecutaRotinaEnable
        
    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
    
    Grid_Limpa = SUCESSO
                
End Function

Function Grid_Trata_Tecla_Campo(KeyAscii As Integer, objGridInt As AdmGrid) As Long
'faz o tratamento das teclas para celulas de um grid

Dim lErro As Long
Dim iCriticaCampo As Integer

On Error GoTo Erro_Grid_Trata_Tecla_Campo
     
     Call Grid_Trata_Tecla_Campo1(KeyAscii, objGridInt, iCriticaCampo)

     If iCriticaCampo = 1 Then

        'critica a saida da celula
        lErro = objGridInt.objForm.Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 20196

        If objGridInt.iNaoTrataEnter <> MARCADO Then
            Call Grid_Trata_Tecla_Campo2(KeyAscii, objGridInt)
        End If

        'reseta a indicação que teclou escape ou enter
        objGridInt.iEscEnter = 0
            

    End If
    
    objGridInt.iNaoTrataEnter = DESMARCADO

    Grid_Trata_Tecla_Campo = SUCESSO

    Exit Function

Erro_Grid_Trata_Tecla_Campo:

    objGridInt.iNaoTrataEnter = DESMARCADO

    Grid_Trata_Tecla_Campo = Err

    Select Case Err

        Case 20196
'subtituido por mario em 4/8/99 pois quando dava erro nao estava zerando lErroSaidaCelula
'             Call Grid_Erro_Trata_Tecla_Campo(objGridInt)
             Call Grid_Erro_Trata_Tecla_Campo_Contabil(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142015)

    End Select

    Exit Function

End Function

Function Grid_Trata_Tecla_Campo_Contabil(KeyAscii As Integer, objGridInt As AdmGrid) As Long
'faz o tratamento das teclas para celulas de um grid

Dim lErro As Long
Dim iCriticaCampo As Integer

On Error GoTo Erro_Grid_Trata_Tecla_Campo_Contabil

     Call Grid_Trata_Tecla_Campo1(KeyAscii, objGridInt, iCriticaCampo)

     If iCriticaCampo = 1 Then

         'critica a saida da celula
         lErro = objGridInt.objForm.Saida_Celula(objGridInt)
         If lErro <> SUCESSO Then Error 20644

         Call Grid_Trata_Tecla_Campo2(KeyAscii, objGridInt)

        'reseta a indicação que teclou escape ou enter
        objGridInt.iEscEnter = 0
            

    End If

    Grid_Trata_Tecla_Campo_Contabil = SUCESSO

    Exit Function

Erro_Grid_Trata_Tecla_Campo_Contabil:

    Grid_Trata_Tecla_Campo_Contabil = Err

    Select Case Err

        Case 20644
             Call Grid_Erro_Trata_Tecla_Campo_Contabil(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142016)

    End Select

    Exit Function

End Function

Private Function Grid_Trata_Tecla_Campo1(KeyAscii As Integer, objGridInt As AdmGrid, iCriticaCampo As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Grid_Trata_Tecla_Campo1

    iCriticaCampo = 0

    'se digitou ESCAPE ou ENTER
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then

        'customizacao utilizada na fox para ao dar enter passar de um segmento para outro no campo produto
        lErro = CF("Produto_Segmento", objGridInt, KeyAscii)
        If lErro <> SUCESSO And lErro <> 105117 Then gError 105116

        If lErro = SUCESSO Then
            
            'faz o retangulo de indicação do foco no grid desaparecer
            objGridInt.objGrid.FocusRect = flexFocusNone
    
            'indica que teclou escape ou enter
            objGridInt.iEscEnter = 1

            iCriticaCampo = 1
        
        End If
        
    End If

    Grid_Trata_Tecla_Campo1 = SUCESSO

    Exit Function

Erro_Grid_Trata_Tecla_Campo1:

    Grid_Trata_Tecla_Campo1 = gErr

    Select Case gErr

        Case 105116

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142017)

    End Select

    Exit Function

End Function

Private Function Grid_Trata_Tecla_Campo2(KeyAscii As Integer, objGridInt As AdmGrid) As Long

Dim iProximo_Campo_Habilitado As Integer
Dim objCampo As Object
Dim iColuna As Integer
Dim iLinha As Integer
Dim objCampo1 As Object
Dim iCount As Integer

Const DEBUG_GRAVA_LOG = 0

    'se foi digitado um ENTER
    If KeyAscii = vbKeyReturn Then
    
        If DEBUG_GRAVA_LOG = MARCADO Then
            Call Grava_Log2("========================================================================")
            Call Grava_Log2("INICIO FUNÇÃO: Grid_Trata_Tecla_Campo2")
            Call Grava_Log2("objGridInt: col " & CStr(objGridInt.objGrid.Col) & "; cols " & CStr(objGridInt.objGrid.Cols))
            Call Grava_Log2("objGridInt: row " & CStr(objGridInt.objGrid.Row) & "; rows " & CStr(objGridInt.objGrid.Rows))
            Call Grava_Log2("objGridInt: iLinhasExistentes " & CStr(objGridInt.iLinhasExistentes) & "; iLinhaAntiga " & CStr(objGridInt.iLinhaAntiga))
            Call Grava_Log2(" ")
            'mario foco
            
            'If objGridInt.objGrid.TextMatrix(0, 1) = "PRETO" Then objGridInt.objGrid.TextMatrix(0, 1) = ""
            
            'objGridInt.objGrid.TextMatrix(0, 1) = "9" & objGridInt.objGrid.TextMatrix(0, 1)
        End If
        
        objGridInt.iEscEnter = 0
        
        'coloca o foco no grid
        If objGridInt.objGrid.Visible = True Then objGridInt.objGrid.SetFocus

'        'se está posicionado na ultima coluna
'
'        If objGridInt.objGrid.Cols - 1 = objGridInt.objGrid.Col Then
'
'            'vai para o proximo campo
'            If objGridInt.objGrid.Row = objGridInt.objGrid.Rows - 1 Then
'                SendKeys "^{HOME}{ENTER}"
'            Else
'                SendKeys "{HOME}{DOWN}{ENTER}"
'            End If
'
'        Else
'            'vai para o proximo campo
'            SendKeys "{RIGHT}{ENTER}"
'        End If

' novo
        
        iProximo_Campo_Habilitado = 0
        iColuna = objGridInt.objGrid.Col
        
        '05/03/18
        If iColuna <> 0 Then
            Set objCampo1 = objGridInt.objForm.Controls(objGridInt.colCampo(iColuna))
            If TypeName(objCampo1) = "Object" Then Set objCampo1 = objCampo1(objGridInt.colIndex(iColuna))
        Else
            Set objCampo1 = Nothing
        End If
        
'novo 24/12/99 - para que este rotina funcione plenamente as implementacao de Grid_Rotina_Enable nas telas não podem usar Grid.Row ou Grid.Col
'        iLinha = objGridInt.objGrid.Row
'
'        Do While True
'
'            If objGridInt.objGrid.Cols - 1 = iColuna Then
'
'                If objGridInt.objGrid.Row = objGridInt.objGrid.Rows - 1 Then
'                    iLinha = objGridInt.objGrid.FixedRows
'                Else
'                    iLinha = iLinha + 1
'                End If
'
'                iColuna = objGridInt.objGrid.FixedCols
'
'            Else
'
'                iColuna = iColuna + 1
'
'            End If
'
'            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(iColuna))
'
'            If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
'                Call objGridInt.objForm.Rotina_Grid_Enable(iLinha, objCampo, ROTINA_GRID_TRATA_TECLA_CAMPO2)
'            End If
'
'            If objCampo.Enabled = True Then
'                objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO
'                objGridInt.objGrid.Row = iLinha
'                objGridInt.objGrid.Col = iColuna
'                objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
'                Exit Do
'            End If
'
'        Loop
        
'fim novo 24/12/99
        
        iLinha = objGridInt.objGrid.Row
        
        'se está posicionado na ultima coluna
        If objGridInt.objGrid.Cols - 1 = iColuna Then
        
            'vai para o proximo campo
            If objGridInt.objGrid.Row = objGridInt.objGrid.Rows - 1 Then
                If DEBUG_GRAVA_LOG = MARCADO Then Call Grava_Log2("01 SendKeys '^{HOME}', True")
                SendKeys "^{HOME}", True
            Else
                If DEBUG_GRAVA_LOG = MARCADO Then Call Grava_Log2("02 SendKeys '{HOME}{DOWN}'")
                SendKeys "{HOME}{DOWN}", True
                If objGridInt.iLinhasExistentes <= objGridInt.objGrid.Row Then iLinha = iLinha + 1
            End If

            iColuna = objGridInt.objGrid.FixedRows

        Else
            'vai para o proximo campo
            SendKeys "{RIGHT}", True

            iColuna = iColuna + 1

        End If
        
        iCount = 0

        If DEBUG_GRAVA_LOG = MARCADO Then
            Call Grava_Log2("INÍCIO: Do While iProximo_Campo_Habilitado = 0")
            Call Grava_Log2(" ")
        End If
        
        Do While iProximo_Campo_Habilitado = 0
        
            iCount = iCount + 1
            If DEBUG_GRAVA_LOG = MARCADO Then Call Grava_Log2("FUNÇÃO: Grid_Trata_Tecla_Campo2 COUNT(iProximo_Campo_Habilitado) = " & CStr(iCount))
            
            'If iCount > 100 Then Exit Do

            Set objCampo = objGridInt.objForm.Controls(objGridInt.colCampo(iColuna))

            If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iColuna))

            If objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE Then
  '              MsgBox ("RGE objGridInt.objGrid.Row = " & objGridInt.objGrid.Row)

                Call objGridInt.objForm.ROTINA_GRID_ENABLE(iLinha, objCampo, ROTINA_GRID_TRATA_TECLA_CAMPO2)
            End If

            If objCampo.Enabled = True Then

                iProximo_Campo_Habilitado = 1

            Else

                'se a coluna atual é a ultima coluna
                If objGridInt.objGrid.Cols - 1 = iColuna Then

                    'vai para o proximo campo
                    If objGridInt.objGrid.Row = objGridInt.objGrid.Rows - 1 Then
                        If DEBUG_GRAVA_LOG = MARCADO Then
                            Call Grava_Log2("03 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                            Call Grava_Log2("03 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                            Call Grava_Log2("03 SendKeys '^{HOME}', True")
                        End If
                        SendKeys "^{HOME}", True
                        iLinha = 1
                        If DEBUG_GRAVA_LOG = MARCADO Then
                            Call Grava_Log2("03 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                            Call Grava_Log2("03 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                        End If
                    Else
                        If DEBUG_GRAVA_LOG = MARCADO Then
                            Call Grava_Log2("04 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                            Call Grava_Log2("04 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                            Call Grava_Log2("04 SendKeys '{HOME}{DOWN}'")
                        End If
                        SendKeys "{HOME}{DOWN}", True
                        If objGridInt.iLinhasExistentes <= objGridInt.objGrid.Row Then iLinha = iLinha + 1
                        If DEBUG_GRAVA_LOG = MARCADO Then
                            Call Grava_Log2("04 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                            Call Grava_Log2("04 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                        End If
                    End If

                    iColuna = objGridInt.objGrid.FixedCols

                Else
                    If DEBUG_GRAVA_LOG = MARCADO Then
                        Call Grava_Log2("05 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                        Call Grava_Log2("05 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                        Call Grava_Log2("05 SendKeys '{RIGHT}', True")
                    End If
                    'vai para o proximo campo
                    SendKeys "{RIGHT}", True

                    iColuna = iColuna + 1

                    If DEBUG_GRAVA_LOG = MARCADO Then
                        Call Grava_Log2("05 iLinha = " & CStr(iLinha) & " iColuna = " & CStr(iColuna))
                        Call Grava_Log2("05 objGridInt.row " & CStr(objGridInt.objGrid.Row) & "; objGridInt.col " & CStr(objGridInt.objGrid.Col))
                    End If
                End If

            End If

        Loop
        
        If DEBUG_GRAVA_LOG = MARCADO Then
            Call Grava_Log2(" ")
            Call Grava_Log2("FIM: Do While iProximo_Campo_Habilitado = 0")
            Call Grava_Log2(" ")
        End If
        
        'novo
        'quando estiver lidando com uma checkbox ou um optionbutton não pode enviar um ENTER
        'Se enviar vai passar para o proximo campo
        If Controle_ObterNomeClasse(objCampo) = "CheckBox" Or Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
            objGridInt.objGrid.FocusRect = flexFocusHeavy
        '05/03/18
        'ElseIf objCampo1.Enabled Then
        '    SendKeys "{ENTER}", True
        ElseIf Not (objCampo1 Is Nothing) Then
            If objCampo1.Enabled Then SendKeys "{ENTER}", True
        Else
            SendKeys "{ENTER}", True
        End If
'Fim novo

    End If
        
    
    If KeyAscii = vbKeyEscape Then
        
        'faz o retangulo de indicação do foco no grid aparecer
        objGridInt.objGrid.FocusRect = flexFocusHeavy

        'coloca o foco no grid
        If objGridInt.objGrid.Visible = True Then objGridInt.objGrid.SetFocus

    End If

    'não deixa o caracter digitado prosseguir no seu processamento
    KeyAscii = 0

    'objGridInt.objGrid.TextMatrix(0, 1) = "5" & objGridInt.objGrid.TextMatrix(0, 1)

    If DEBUG_GRAVA_LOG = MARCADO Then
        Call Grava_Log2("INICIO FUNÇÃO: Grid_Trata_Tecla_Campo2")
        Call Grava_Log2("========================================================================")
        Call Grava_Log2(" ")
    End If

    Grid_Trata_Tecla_Campo2 = SUCESSO

End Function

Function Grid_Erro_Trata_Tecla_Campo(objGridInt As AdmGrid) As Long

    If objGridInt.lErroSaidaCelula = 3 Then objGridInt.lErroSaidaCelula = 0
    
    If objGridInt.lErroSaidaCelula = 2 Then objGridInt.lErroSaidaCelula = 3
    
    Grid_Erro_Trata_Tecla_Campo = SUCESSO
        
End Function

Function Grid_Erro_Trata_Tecla_Campo_Contabil(objGridInt As AdmGrid) As Long

    If objGridInt.lErroSaidaCelula = 3 Or objGridInt.lErroSaidaCelula = 1 Then
        
        'adicionado por mario em 21/9/99 em função do MDI
        If objGridInt.lErroSaidaCelula = 1 Then
            If objGridInt.objControle.Visible = True Then objGridInt.objControle.SetFocus
        End If
        
        objGridInt.lErroSaidaCelula = 0
        
    End If
    
    If objGridInt.lErroSaidaCelula = 2 Then objGridInt.lErroSaidaCelula = 3
    
    Grid_Erro_Trata_Tecla_Campo_Contabil = SUCESSO
        
End Function

Function Grid_Campo_Libera_Foco(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCriticaCampo As Integer
Dim objForm As Form, objFormAtivo As Object
Dim iIndice As Integer

On Error GoTo Erro_Grid_Campo_Libera_Foco

''    'mario colocado devido ao MDI em 23/9/99
''    If Not (GL_objMDIForm Is Nothing) Then
''        If Not (objGridInt.objForm Is GL_objMDIForm.ActiveForm.objFormOcx) Then
''            objGridInt.objForm.Show
''        End If
''    End If
''
    'Jones alterou o codigo acima pois dava erro no dic em 15/11/99
    If Not (GL_objMDIForm Is Nothing) Then
        Set objFormAtivo = GL_objMDIForm.ActiveForm
        If Not (objFormAtivo Is Nothing) Then
            'se o form contiver apenas um usercontrol (form em ocx)
            If objFormAtivo.Controls.Count = 1 Then
                If objGridInt.objForm.Name <> objFormAtivo.objFormOcx.Name Then objGridInt.objForm.Show 'obs.: nao pude testar por 'Is' devido aos usercontrols em que objGridInt.objForm é nao verdade um objeto de classe ct*
            Else
                If Not (objGridInt.objForm Is objFormAtivo) Then objGridInt.objForm.Show
            End If
        End If
    End If

    'mario colocado devido ao MDI em 23/9/99
    If objGridInt.iSaidaCelula = 4 Then objGridInt.iSaidaCelula = 1

    lErro = Grid_Campo_Libera_Foco1(objGridInt, iCriticaCampo)
    If lErro <> SUCESSO Then Error 20197

    If iCriticaCampo = 1 Then

        lErro = objGridInt.objForm.Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 20198

        Call Grid_Campo_Libera_Foco2(objGridInt)

    End If

    Grid_Campo_Libera_Foco = SUCESSO

    Exit Function

Erro_Grid_Campo_Libera_Foco:

    Grid_Campo_Libera_Foco = Err

    Select Case Err

        Case 20197

        Case 20198
            Call Grid_Retorna_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142018)

    End Select

    Exit Function

End Function

Function Grid_Campo_Libera_Foco_Modal(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCriticaCampo As Integer
Dim objForm As Form
Dim iIndice As Integer

On Error GoTo Erro_Grid_Campo_Libera_Foco_Modal

    'mario colocado devido ao MDI em 23/9/99
    If objGridInt.iSaidaCelula = 4 Then objGridInt.iSaidaCelula = 1

    lErro = Grid_Campo_Libera_Foco1(objGridInt, iCriticaCampo)
    If lErro <> SUCESSO Then Error 55775

    If iCriticaCampo = 1 Then

        lErro = objGridInt.objForm.Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 55776

        Call Grid_Campo_Libera_Foco2(objGridInt)

    End If

    Grid_Campo_Libera_Foco_Modal = SUCESSO

    Exit Function

Erro_Grid_Campo_Libera_Foco_Modal:

    Grid_Campo_Libera_Foco_Modal = Err

    Select Case Err

        Case 55775

        Case 55776
            Call Grid_Retorna_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142019)

    End Select

    Exit Function

End Function

Function Grid_Campo_Libera_Foco_Contabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCriticaCampo As Integer

On Error GoTo Erro_Grid_Campo_Libera_Foco_Contabil

    'mario colocado devido ao MDI em 23/9/99
    If Not (GL_objMDIForm Is Nothing) Then
        If Not (objGridInt.objForm Is GL_objMDIForm.ActiveForm.objFormOcx) Then
            objGridInt.objForm.Show
        End If
    End If

    'mario novo
    If objGridInt.iSaidaCelula = 4 Then objGridInt.iSaidaCelula = 1

    lErro = Grid_Campo_Libera_Foco1(objGridInt, iCriticaCampo)
    If lErro <> SUCESSO Then Error 20197

    If iCriticaCampo = 1 Then

        lErro = objGridInt.objForm.Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 20198

        Call Grid_Campo_Libera_Foco2(objGridInt)

    End If

    Grid_Campo_Libera_Foco_Contabil = SUCESSO

    Exit Function

Erro_Grid_Campo_Libera_Foco_Contabil:

    Grid_Campo_Libera_Foco_Contabil = Err

    Select Case Err

        Case 20197

        Case 20198
            Call Grid_Retorna_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142020)

    End Select

    Exit Function

End Function


Function Grid_Refresh_Checkbox(objGridInt As AdmGrid) As Long
'faz refresh nas colunas de checkbo e optionbutton

Dim iColAtual As Integer
Dim iLinhaAtual As Integer
Dim iCol As Integer
Dim iLinha As Integer
Dim vName As Variant
Dim objCampo As Object, objControls As Object
Dim iIndice As Integer, objGrid As Object
Dim objCheckboxUnchecked As Object, objCheckboxChecked As Object
Dim objOptionButtonUnchecked As Object, objOptionButtonChecked As Object
Dim objCheckboxGrayed As Object

    Set objGrid = objGridInt.objGrid
    
    'desabilita temporariamente o redesenho
    objGrid.Redraw = False

    'guarda a linha e coluna atuais
    iColAtual = objGrid.Col
    iLinhaAtual = objGrid.Row
    
    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO
    
    iCol = 0

    Set objControls = objGridInt.objForm.Controls
    
    For Each vName In objGridInt.colCampo
        iCol = iCol + 1
        Set objCampo = objControls(vName)
    
        If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iCol))
    
        If Controle_ObterNomeClasse(objCampo) = "CheckBox" Then
                        
            With objGrid
                
                Set objCheckboxUnchecked = objGridInt.objCheckboxUnchecked
                Set objCheckboxChecked = objGridInt.objCheckboxChecked
                Set objCheckboxGrayed = objGridInt.objCheckboxGrayed
                
                .Col = iCol
                For iIndice = 1 To .Rows - 1
                    .Row = iIndice
                    
                    If .Text = "0" Or Len(.Text) = 0 Then
                        Set .CellPicture = objCheckboxUnchecked
                    ElseIf .Text = "2" Then
                        Set .CellPicture = objCheckboxGrayed
                    Else
                        Set .CellPicture = objCheckboxChecked
                    End If
                Next
            End With
        End If
            
        If Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
                        
            With objGrid
            
                Set objOptionButtonUnchecked = objGridInt.objOptionButtonUnchecked
                Set objOptionButtonChecked = objGridInt.objOptionButtonChecked
                
                .Col = iCol
                For iIndice = 1 To .Rows - 1
                    .Row = iIndice
                    
                    If .Text = "0" Or Len(.Text) = 0 Then
                        Set .CellPicture = objOptionButtonUnchecked
                    Else
                        Set .CellPicture = objOptionButtonChecked
                    End If
                Next
            End With
        End If
            
    Next
    
    'retorna a linha e coluna atuais
    objGrid.Row = iLinhaAtual
    objGrid.Col = iColAtual
            
    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
            
    'reabilita o redesenho
    objGrid.Redraw = True
    
    Grid_Refresh_Checkbox = SUCESSO

End Function

Public Sub ListBox_Select(ByVal sString As String, objListBox As Object)
Attribute ListBox_Select.VB_Description = "Seleciona  sString na objListBox, se existir.  objListBox pode ser uma listbox ou combobox."
'Seleciona ítem sString na ListBox/ComboBox, se existir
    
Dim iIndice As Integer
    
    For iIndice = 0 To objListBox.ListCount - 1
        
        If objListBox.List(iIndice) = sString Then
            
            objListBox.ListIndex = iIndice
            Exit For
        
        End If
        
    Next

End Sub

'Function Obter_Usuario(sCodUsuario As String) As Long
''obtem o usuario atual
'
'Dim lErro As Long
'
'On Error GoTo Erro_Obter_Usuario
'
''    lErro = Obter_UsuarioInt(sCodUsuario)
''    If lErro <> SUCESSO Then Error 9038
'
''    sCodUsuario = StringZ(sCodUsuario)
'
'    sCodUsuario = "dudu"
'
'    Obter_Usuario = SUCESSO
'
'    Exit Function
'
'Erro_Obter_Usuario:
'
'    Obter_Usuario = Err
'
'    Select Case Err
'
'        Case 9038
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142021)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function Obter_Modulo(sModulo As String) As Long
'    ''''???? NAO IMPLEMENTADA
'    sModulo = "CTB"
'End Function

Function Obter_Grupo(sCodGrupo As String) As Long
'obtem o grupo atual

Dim lErro As Long

On Error GoTo Erro_Obter_Grupo

    sCodGrupo = String(STRING_GRUPO, 0)
    
    lErro = Obter_GrupoInt(GL_lAcesso, sCodGrupo)
    If lErro <> AD_BOOL_TRUE Then Error 9074

    sCodGrupo = StringZ(sCodGrupo)

    Obter_Grupo = SUCESSO

    Exit Function

Erro_Obter_Grupo:

    Obter_Grupo = Err

    Select Case Err

        Case 9074

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142022)

    End Select

    Exit Function

End Function

Function Sistema_ExecutarRel(lpNomeTsk As String, lpRelOp As Long, ByVal iPrevia As Integer, iOutOfProcess As Integer) As Long
'retorna AD_BOOL_TRUE p/sucesso
Dim lhWnd As Long
 
    If GL_objMDIForm Is Nothing Then
        lhWnd = 0
    Else
        lhWnd = GL_objMDIForm.hWnd
    End If
 
    Sistema_ExecutarRel = Sistema_ExecutarRelInt(GL_lSistema, lpNomeTsk, lpRelOp, iPrevia, gdtDataAtual, lhWnd, iOutOfProcess)
    
End Function

Function Sist_ImpressoraDlg(ByVal bConfig As Integer, Optional iLandscape As Integer = 0) As Long
'retorna AD_BOOL_TRUE p/sucesso
    Sist_ImpressoraDlg = Sist_ImpressoraDlgInt(GL_lSistema, bConfig, iLandscape)
End Function

Function Sistema_EditarRel(lpNomeTsk As String) As Long
'retorna AD_BOOL_TRUE p/sucesso
    Sistema_EditarRel = Sistema_EditarRelInt(GL_lSistema, lpNomeTsk)
End Function

Function Teste_Salva(objTela As Object, iAlterado As Integer) As Long

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim asReg(1 To 100) As String
Dim iIndice As Integer

On Error GoTo Erro_Teste_Salva

iIndice = 1

asReg(iIndice) = "****************************************************************************"
iIndice = iIndice + 1
asReg(iIndice) = Format(Now, "dd/mm/yyyy hh:mm:ss") & "    -    " & gsNomeEmpresa & "    -    " & gsUsuario
iIndice = iIndice + 1
asReg(iIndice) = "AdmLib.ADM.Teste_Salva "
iIndice = iIndice + 1
asReg(iIndice) = " "
iIndice = iIndice + 1

asReg(iIndice) = "1 - iAlterado = " & CStr(iAlterado)
iIndice = iIndice + 1

    If iAlterado = REGISTRO_ALTERADO Then

asReg(iIndice) = "2 - Rotina_Aviso(AVISO_DESEJA_SALVAR_ALTERACOES) "
iIndice = iIndice + 1

        vbMsgRes = Rotina_Aviso(vbYesNoCancel, "AVISO_DESEJA_SALVAR_ALTERACOES")

asReg(iIndice) = "3 - vbMsgRes (Rotina_Aviso(AVISO_DESEJA_SALVAR_ALTERACOES)) =  " & CStr(vbMsgRes)
iIndice = iIndice + 1

        If vbMsgRes = vbYes Then
            
asReg(iIndice) = "4 - objTela.Gravar_Registro"
iIndice = iIndice + 1

            lErro = objTela.Gravar_Registro()
            
asReg(iIndice) = "5 - lErr (objTela.Gravar_Registro) =  " & CStr(lErro)
iIndice = iIndice + 1

            If lErro <> SUCESSO Then Error 20322

        ElseIf vbMsgRes = vbCancel Then

asReg(iIndice) = "6 - Error 20323  "
iIndice = iIndice + 1

            Error 20323

        End If

        iAlterado = 0

    End If

    Teste_Salva = SUCESSO

    Exit Function

Erro_Teste_Salva:

    Teste_Salva = Err

asReg(iIndice) = "7 - gErr = " & CStr(gErr) & " Error = " & Error
iIndice = iIndice + 1

asReg(iIndice) = "****************************************************************************"
iIndice = iIndice + 1

   
    Select Case Err

        Case 20322, 20323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142023)

    End Select

    'Call Grava_Log(asReg)

    Exit Function

End Function

Sub Tela_QueryUnload(objTela As Object, iAlterado As Integer, iCancel As Integer, iUnloadMode As Integer, Optional iTelaCorrenteAtiva As Integer = 1)

Dim lErro As Long
Dim objControl As Object

    If iTelaCorrenteAtiva <> 0 Then

'           comentado Jones 06/05/2000
'''        If iUnloadMode <> vbFormCode Then
'''
'''            Set objControl = objTela.ActiveControl
'''
'''            SendKeys "{TAB}", True
'''
'''            If objControl.Name = objTela.ActiveControl.Name Then
'''                iCancel = 1
'''                Exit Sub
'''            End If
'''
'''        End If

        lErro = Teste_Salva(objTela, iAlterado)
        If lErro <> SUCESSO Then iCancel = 1

    End If
            
End Sub

'Formata quantidade com formato de ESTOQUE.
'Está nesse nível porque é usada pelo AdmGrid.
'Os controles do Grid com quantidades devem ter .format
'incializado com FORMATO_ESTOQUE
Function Formata_Estoque(dQuantidade As Double) As String

    Formata_Estoque = Format(dQuantidade, FORMATO_ESTOQUE)

End Function

'Formata custos com formato de custo.
'Está nesse nível porque é usada pelo AdmGrid.
'Os controles do Grid com quantidades devem ter .format
'incializado com FORMATO_CUSTO
Function Formata_Custo(dQuantidade As Double) As String

    Formata_Custo = Format(dQuantidade, FORMATO_CUSTO)

End Function

Function Resto(dDividendo As Double, iDivisor As Integer) As Integer
'retorna o resto (inteiro) da divisao de dDividendo por iDivisor

Dim dQuociente As Double
Dim iResto As Integer
Dim iPosicaoVirgula As Integer

On Error GoTo Erro_Resto
    
'    If dDividendo >= iDivisor Then
'        dQuociente = dDividendo / iDivisor
'        dQuociente = Int(dQuociente)
'        iResto = CInt(dDividendo - (iDivisor * dQuociente))
'    Else
'        iResto = CInt(dDividendo)
'
'    End If
    
    If dDividendo >= iDivisor Then
        dQuociente = dDividendo / iDivisor
        iPosicaoVirgula = InStr(1, CStr(dQuociente), ",")
        If iPosicaoVirgula = 0 Then
            iResto = 0
        Else
            dQuociente = CDbl(Mid(CStr(dQuociente), 1, iPosicaoVirgula - 1))
            iResto = CInt(dDividendo - (iDivisor * dQuociente))
        End If
    Else
        iResto = CInt(dDividendo)
    
    End If
    
    Resto = iResto
    
    Exit Function
    
Erro_Resto:

    Resto = 0
    
    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142024)
    
    Exit Function
    
End Function

Private Function Rotina_ObterFuncaoVB(ByVal sCodigoRotina As String) As Long

Dim lErro As Long, X As Object, sSigla As String, lTransacao As Long
Dim lComando As Long, Y As New AdmSQL, sCodGrupo As String, sProjClass As String, lComando2 As Long
Dim sGrupoRotinasProjeto As String, sGrupoRotinasClasse As String, sRotinasProjetoOriginal As String, sRotinasClasseOriginal As String, sRotinasProjetoCustomizado As String, sRotinasClasseCustomizado As String

On Error GoTo Erro_Rotina_ObterFuncaoVB

    If gcolRotinasSGE.Exists(CVar(UCase(sCodigoRotina))) Then
    
        mvarvBufferObterRot = gcolRotinasSGE(CVar(UCase(sCodigoRotina)))
    
    Else
    
        If gcolRotinasSGE.Count = 3 Then
        
            Call Obter_Grupo(sCodGrupo)
            
            lComando = Y.Comando_AbrirExt(GL_lConexaoDicBrowse)
            If lComando = 0 Then Error 59277
            
            sGrupoRotinasProjeto = String(NOME_PROJETO, 0)
            sGrupoRotinasClasse = String(NOME_CLASSE, 0)
            sRotinasProjetoOriginal = String(NOME_PROJETO, 0)
            sRotinasClasseOriginal = String(NOME_CLASSE, 0)
            sRotinasProjetoCustomizado = String(NOME_PROJETO, 0)
            sRotinasClasseCustomizado = String(NOME_CLASSE, 0)
            sSigla = String(255, 0)
            
            lErro = Y.Comando_Executar(lComando, "SELECT Rotinas.Sigla, GrupoRotinas.Projeto, GrupoRotinas.Classe, Rotinas.Projeto_original, Rotinas.Classe_original, Rotinas.Projeto_customizado, Rotinas.Classe_customizado FROM GrupoRotinas, Rotinas WHERE GrupoRotinas.CodGrupo=? AND GrupoRotinas.SiglaRotina = Rotinas.Sigla AND TipoDeAcesso = 1 AND Rotinas.PreLoad = 1", _
                sSigla, sGrupoRotinasProjeto, sGrupoRotinasClasse, sRotinasProjetoOriginal, sRotinasClasseOriginal, sRotinasProjetoCustomizado, sRotinasClasseCustomizado, sCodGrupo)
            If lErro <> AD_SQL_SUCESSO Then Error 59278
            
            lErro = Y.Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59279
            
            Do While lErro = AD_SQL_SUCESSO
            
                If sGrupoRotinasProjeto = "" Then
                    If sRotinasProjetoCustomizado <> "" Then
                        sProjClass = sRotinasProjetoCustomizado & "." & sRotinasClasseCustomizado
                    Else
                        sProjClass = sRotinasProjetoOriginal & "." & sRotinasClasseOriginal
                    End If
                Else
                    sProjClass = sGrupoRotinasProjeto & "." & sGrupoRotinasClasse
                End If
                
                gcolRotinasSGE.Add UCase(sSigla), sProjClass
                
                lErro = Y.Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59280
            
            Loop
            
            Call Y.Comando_Fechar(lComando)
        
        End If
    
        If gcolRotinasSGE.Exists(CVar(UCase(sCodigoRotina))) Then
        
            mvarvBufferObterRot = gcolRotinasSGE(CVar(UCase(sCodigoRotina)))
        
        Else
    
            Call Obter_Grupo(sCodGrupo)
            
            lComando = Y.Comando_AbrirExt(GL_lConexaoDicBrowse)
            If lComando = 0 Then Error 59277
            
            lComando2 = Y.Comando_AbrirExt(GL_lConexaoDic)
            If lComando2 = 0 Then Error 59277
            
            sGrupoRotinasProjeto = String(NOME_PROJETO, 0)
            sGrupoRotinasClasse = String(NOME_CLASSE, 0)
            sRotinasProjetoOriginal = String(NOME_PROJETO, 0)
            sRotinasClasseOriginal = String(NOME_CLASSE, 0)
            sRotinasProjetoCustomizado = String(NOME_PROJETO, 0)
            sRotinasClasseCustomizado = String(NOME_CLASSE, 0)
            
            lErro = Y.Comando_Executar(lComando, "SELECT GrupoRotinas.Projeto, GrupoRotinas.Classe, Rotinas.Projeto_original, Rotinas.Classe_original, Rotinas.Projeto_customizado, Rotinas.Classe_customizado FROM GrupoRotinas, Rotinas WHERE GrupoRotinas.CodGrupo=? AND GrupoRotinas.SiglaRotina = Rotinas.Sigla AND TipoDeAcesso = 1 AND Rotinas.Sigla = ?", _
                sGrupoRotinasProjeto, sGrupoRotinasClasse, sRotinasProjetoOriginal, sRotinasClasseOriginal, sRotinasProjetoCustomizado, sRotinasClasseCustomizado, sCodGrupo, sCodigoRotina)
            If lErro <> AD_SQL_SUCESSO Then Error 59278
            
            lErro = Y.Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59279
            
            If lErro = AD_SQL_SUCESSO Then
            
                If sGrupoRotinasProjeto = "" Then
                    If sRotinasProjetoCustomizado <> "" Then
                        sProjClass = sRotinasProjetoCustomizado & "." & sRotinasClasseCustomizado
                    Else
                        sProjClass = sRotinasProjetoOriginal & "." & sRotinasClasseOriginal
                    End If
                Else
                    sProjClass = sGrupoRotinasProjeto & "." & sGrupoRotinasClasse
                End If
                
                gcolRotinasSGE.Add UCase(sCodigoRotina), sProjClass
                
                mvarvBufferObterRot = CVar(sProjClass)
            
                If gbPreLoadGravar Then
                
                    lTransacao = Y.Transacao_AbrirDic
                    If lTransacao = 0 Then Error 59282
                
                    lErro = Y.Comando_Executar(lComando2, "UPDATE Rotinas SET PreLoad = 1 WHERE Sigla = ?", sCodigoRotina)
                    If lErro <> AD_SQL_SUCESSO Then Error 59284
                
                    lErro = Y.Transacao_CommitDic
                    If lErro <> AD_SQL_SUCESSO Then Error 59283
                
                End If
                
            Else
            
                Error 59281
                
            End If
            
            Call Y.Comando_Fechar(lComando)
            Call Y.Comando_Fechar(lComando2)
        
        End If
    
    End If
    
    Rotina_ObterFuncaoVB = SUCESSO
     
    Exit Function
    
Erro_Rotina_ObterFuncaoVB:

    Rotina_ObterFuncaoVB = Err
     
    Select Case Err
          
        Case 59277
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 59278, 59279, 59280
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIC_ROTINAS", Err)
        
        Case 59281
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error & " Erro na leitura da rotina " & sCodigoRotina, 142026)
     
    End Select
     
    If lComando <> 0 Then Call Y.Comando_Fechar(lComando)
    If lComando2 <> 0 Then Call Y.Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function Rotina_ObterObjetoFuncao(sCodigoRotina As String, X As Object) As Long
'Retorna um objeto que implemente uma funcao que tem o local de sua implementacao e a autorizacao de acesso identificados por sCodigoRotina

Dim sProjeto As String, sClasse As String, lErro As Long, sTemp As String

On Error GoTo Erro_Rotina_ObterObjetoFuncao

    If Not gbVPN Then
    
        lErro = Rotina_ObterFuncaoInt(GL_lDicDados, sCodigoRotina, mvarvBufferObterRot)
        If (lErro <> AD_BOOL_TRUE) Then Error 59258
            
    Else
    
        lErro = Rotina_ObterFuncaoVB(sCodigoRotina)
        If lErro <> SUCESSO Then Error 59258
    
    End If
    
    'pesquisa se o objeto foi criado anteriormente e inserido na colecao
    If mvarColObjFunc.Exists(mvarvBufferObterRot) Then
        
        Set X = mvarColObjFunc(mvarvBufferObterRot)
    
    Else
    
        Set X = CreateObject(mvarvBufferObterRot)

        'se nao for p/batch, deixar carregado
        If InStr(1, mvarvBufferObterRot, "Batch", vbTextCompare) = 0 Then Call mvarColObjFunc.Add(mvarvBufferObterRot, X)
        
    End If
    
    Rotina_ObterObjetoFuncao = SUCESSO
     
    Exit Function
    
Erro_Rotina_ObterObjetoFuncao:

    Rotina_ObterObjetoFuncao = Err
     
    Select Case Err
          
        Case 59258
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL1", Err, sCodigoRotina)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error & " : " & sCodigoRotina, 142025)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Initialize()

    If gsLogonId = "" Then gsLogonId = Format(Now, "YYYYMMDDHHMMSS")
    
    mvarvBufferObterRot = String$(NOME_PROJETO + 1 + NOME_CLASSE + 1, 0)
    
    Set mvarColObjFunc = CreateObject("Scripting.Dictionary")
    
    gbTelaReordenando = False
    
    If glADMCount = 0 Then
            
        Set gcolRotinasSGE = CreateObject("Scripting.Dictionary")
        
        gcolRotinasSGE.Add UCase("Modulos_Le_Todos"), "RotinasAdm.AdmSelect"
        gcolRotinasSGE.Add UCase("Tamanhos_Le"), "RotinasAdm.AdmSelect"
        gcolRotinasSGE.Add UCase("Tamanhos_Grava"), "RotinasAdm.AdmGrava"
    
        gbVPN = True
        gbPreLoadGravar = False
        
        glADMCount = 1
        
    Else
    
        glADMCount = glADMCount + 1
        
    End If

End Sub

Private Sub Class_Terminate()

    Set mvarColObjFunc = Nothing
    
    glADMCount = glADMCount - 1
    If glADMCount = 0 Then Set gcolRotinasSGE = Nothing
    
End Sub

Function Grid_Refresh_Checkbox_Limpa(objGridInt As AdmGrid) As Long
'faz refresh nas colunas de checkbo e optionbutton

Dim iColAtual As Integer
Dim iLinhaAtual As Integer
Dim iCol As Integer
Dim iLinha As Integer
Dim vName As Variant
Dim objCampo As Object, objControls As Object
Dim iIndice As Integer, objGrid As Object

    Set objGrid = objGridInt.objGrid
    
    'desabilita temporariamente o redesenho
    objGrid.Redraw = False

    objGridInt.iExecutarFuncao = GRID_NAO_EXECUTAR_FUNCAO
    
    iCol = 0

    Set objControls = objGridInt.objForm.Controls
    
    For Each vName In objGridInt.colCampo
    
        iCol = iCol + 1
        Set objCampo = objControls(vName)
    
        If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridInt.colIndex(iCol))
    
        If Controle_ObterNomeClasse(objCampo) = "CheckBox" Then
                        
            With objGrid
                
                .Col = iCol
                .Row = 1
                .RowSel = .Rows - 1
                .FillStyle = flexFillRepeat
                Set .CellPicture = objGridInt.objCheckboxUnchecked
                .Text = "0"
                
                If vName = "CTBGerencial" And objGridInt.iGerencial = 2 Then
                    Set .CellPicture = objGridInt.objCheckboxGrayed
                    .Text = objGridInt.iGerencial
                End If
                
                .FillStyle = flexFillSingle
                
            End With
            
        End If
            
        If Controle_ObterNomeClasse(objCampo) = "OptionButton" Then
                        
            With objGrid
            
                .Col = iCol
                .Row = 1
                .RowSel = .Rows - 1
                .FillStyle = flexFillRepeat
                Set .CellPicture = objGridInt.objOptionButtonUnchecked
                .FillStyle = flexFillSingle
                
            End With
            
        End If
            
    Next
    
    objGridInt.iExecutarFuncao = GRID_EXECUTAR_FUNCAO
            
    'reabilita o redesenho
    objGrid.Redraw = True
    
    Grid_Refresh_Checkbox_Limpa = SUCESSO

End Function

Public Function Rotinas_Pre_Carga(Optional objColObjFuncNovo As Object, Optional objColObjFuncAnterior As Object) As Long
'cria objetos rotinas*.class* e os mantem carregados para que as chamadas posteriores sejam feitas mais rapidamente

Dim lErro As Long, X As Object
Dim lComando As Long, Y As New AdmSQL, sCodGrupo As String, sProjClass As String
Dim sGrupoRotinasProjeto As String, sGrupoRotinasClasse As String, sRotinasProjetoOriginal As String, sRotinasClasseOriginal As String, sRotinasProjetoCustomizado As String, sRotinasClasseCustomizado As String

On Error GoTo Erro_Rotinas_Pre_Carga

    If objColObjFuncNovo Is Nothing Then Set objColObjFuncNovo = mvarColObjFunc
    
    Call Obter_Grupo(sCodGrupo)
    
    lComando = Y.Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 59277
    
    sGrupoRotinasProjeto = String(NOME_PROJETO, 0)
    sGrupoRotinasClasse = String(NOME_CLASSE, 0)
    sRotinasProjetoOriginal = String(NOME_PROJETO, 0)
    sRotinasClasseOriginal = String(NOME_CLASSE, 0)
    sRotinasProjetoCustomizado = String(NOME_PROJETO, 0)
    sRotinasClasseCustomizado = String(NOME_CLASSE, 0)
    
    lErro = Y.Comando_Executar(lComando, "SELECT DISTINCT GrupoRotinas.Projeto, GrupoRotinas.Classe, Rotinas.Projeto_original, Rotinas.Classe_original, Rotinas.Projeto_customizado, Rotinas.Classe_customizado FROM GrupoRotinas, Rotinas WHERE GrupoRotinas.CodGrupo=? AND GrupoRotinas.SiglaRotina = Rotinas.Sigla AND TipoDeAcesso = 1", _
        sGrupoRotinasProjeto, sGrupoRotinasClasse, sRotinasProjetoOriginal, sRotinasClasseOriginal, sRotinasProjetoCustomizado, sRotinasClasseCustomizado, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 59278
    
    lErro = Y.Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59279
    
    Do While lErro = AD_SQL_SUCESSO
    
        If sGrupoRotinasProjeto = "" Then
            If sRotinasProjetoCustomizado <> "" Then
                sProjClass = sRotinasProjetoCustomizado & "." & sRotinasClasseCustomizado
            Else
                sProjClass = sRotinasProjetoOriginal & "." & sRotinasClasseOriginal
            End If
        Else
            sProjClass = sGrupoRotinasProjeto & "." & sGrupoRotinasClasse
        End If
        
        If sProjClass <> "." And InStr(1, sProjClass, "Batch", vbTextCompare) = 0 Then
        
            If objColObjFuncNovo.Exists(sProjClass) <> True Then
        
                Set X = Nothing
                
                If Not (objColObjFuncAnterior Is Nothing) Then
                
                    If objColObjFuncAnterior.Exists(sProjClass) Then Set X = objColObjFuncAnterior(sProjClass)
                    
                End If
                
                If X Is Nothing Then Set X = CreateObject(sProjClass)
        
                Call objColObjFuncNovo.Add(sProjClass, X)
        
            End If
            
        End If
        
        lErro = Y.Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59280
    
    Loop
    
    Call Y.Comando_Fechar(lComando)
    
    Rotinas_Pre_Carga = SUCESSO
     
    Exit Function
    
Erro_Rotinas_Pre_Carga:

    Rotinas_Pre_Carga = Err
     
    Select Case Err
          
        Case 59277
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 59278, 59279, 59280
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIC_ROTINAS", Err)
        
        Case 429
            MsgBox ("Erro na criação de objeto identificado por: " & sProjClass)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error & " Erro na criação de objeto identificado por: " & sProjClass, 142026)
     
    End Select
     
    Call Y.Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Rotinas_Recarrega_Objetos() As Long

Dim lErro As Long
Dim objColObjFunc2 As Object

On Error GoTo Erro_Rotinas_Recarrega_Objetos

    'descarrega os objetos atuais e "limpa" o dicionario
    Set objColObjFunc2 = CreateObject("Scripting.Dictionary")
    
    'recarrega os novos objetos
    lErro = Rotinas_Pre_Carga(objColObjFunc2, mvarColObjFunc)
    If lErro <> SUCESSO Then Error 59281
    
    Set mvarColObjFunc = objColObjFunc2
    
    Rotinas_Recarrega_Objetos = SUCESSO
     
    Exit Function
    
Erro_Rotinas_Recarrega_Objetos:

    Rotinas_Recarrega_Objetos = Err
     
    Select Case Err
          
        Case 59281
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142027)
     
    End Select
     
    Exit Function

End Function

'********** Inicio de Edicao Tela ***************

Public Sub Controle_DragDrop(Controle As Object, Source As Object, X As Single, Y As Single)

Dim Container As Object
Dim Container1 As Object
    
On Error GoTo Erro_Controle_DragDrop
    
    If gobjmenuEdicao.Checked = True Then
    
        If Not Controle Is gobjTelaAtiva Then

            Set Container = Controle.Container
            X = X + Controle.left
            Y = Y + Controle.top
            Do While Not (Container Is gobjTelaAtiva)
                X = X + Container.left
                Y = Y + Container.top
                Set Container = Container.Container
            Loop

        End If

        Set Container = Source.Container
        Set Source.Container = gobjTelaAtiva

        If X + gsngDragX + Source.Width + giLeft > gobjTelaAtiva.Width Then
            X = gobjTelaAtiva.Width - Source.Width - gsngDragX - 100 - giLeft
        End If

        If Y + gsngDragY + Source.Height + giTop > gobjTelaAtiva.Height Then
            Y = gobjTelaAtiva.Height - Source.Height - gsngDragY - 100 - giTop
        End If

        If Y + gsngDragY + giTop < 0 Then
            Y = -(gsngDragY + giTop)
        End If

        If X + gsngDragX + giLeft < 0 Then
            X = -(gsngDragX + giLeft)
        End If

        Set Source.Container = Container

        If Controle_ObterNomeClasse(Controle) = "Frame" Or Controle Is gobjTelaAtiva Then
            If Not (Source.Container Is Controle) And Not (Source Is Controle) Then Set Source.Container = Controle
        ElseIf Not Source.Container Is Controle.Container Then
            Set Source.Container = Controle.Container
        End If

        Set Container1 = Source.Container

        Do While Not (Container1 Is gobjTelaAtiva)
            X = X - Container1.left
            Y = Y - Container1.top
            Set Container1 = Container1.Container
        Loop
            
        'se for um maskedbox e estiver sendo posicionado sobre um controle que não ele mesmo. Faz um deslocamento devido ao problema do maskedbox
        If Not (Controle Is Source) And Controle_ObterNomeClasse(Source) = "MaskEdBox" Then
            giLeft = giLeft - 30
            giTop = giTop - 30
        End If
            
        Source.Move X + gsngDragX + giLeft, Y + gsngDragY + giTop
        Call gobjPropriedades.ComboCampos_Seleciona(Source)
        Set gobjControleDrag = Nothing

    End If
    
    Exit Sub
    
Erro_Controle_DragDrop:

    Select Case Err
    
        Case 425
            Resume Next
            
        Case Else
            MsgBox Err
            
    End Select
    
    Exit Sub

End Sub

Public Sub Controle_MouseDown(Controle As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Container As Object
Dim ContainerAux As Object
Dim iContainer As Integer

On Error GoTo Erro_Controle_MouseDown

    If gobjmenuEdicao.Checked = True And giProxButtonUp = 0 Then
        
        Set ContainerAux = Controle.Container
        
        iContainer = 0
        
        'Seta a tela do label como Tela Ativa
        Do While iContainer = 0
            Set ContainerAux = ContainerAux.Container
        Loop
        
        Set gobjTelaAtiva = ContainerAux
    
        Set gobjControleDrag = Controle
        Controle.Drag 1
        
        Call gobjPropriedades.ComboCampos_Seleciona(Controle)
    
        giProxButtonUp = 1
        giLeft = Controle.left
        giTop = Controle.top
        Set Container = Controle.Container
        Do While Not (Container Is gobjTelaAtiva)
            giLeft = giLeft + Container.left
            giTop = giTop + Container.top
            Set Container = Container.Container
        Loop
        gsngDragX = -X - giLeft
        gsngDragY = -Y - giTop
        
    End If

    Exit Sub

Erro_Controle_MouseDown:

    Select Case Err
        
        Case 438
            iContainer = 1
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142028)
     
    End Select
     
    Exit Sub

End Sub

'********** Fim de Edicao Tela ***************

Sub Sistema_Libera_Globais()

    Set gobjCheckboxChecked = Nothing
    Set gobjCheckboxUnchecked = Nothing
    Set gobjCheckboxGrayed = Nothing
    Set gobjOptionButtonChecked = Nothing
    Set gobjOptionButtonUnChecked = Nothing
    Set gobjButton = Nothing
    
    Set gcolModulo = Nothing
    Set GL_objMDIForm = Nothing
    Set gcolCampos = Nothing
    Set gcolTiposMovEst = Nothing
    Set gobjEstInicial = Nothing
    Set gobjTelaAtiva = Nothing
    Set gobjControleDrag = Nothing
    Set gobjControleAlvo = Nothing
    Set gobjmenuEdicao = Nothing
    Set gcolWndProc = Nothing
    Set gobjPropriedades = Nothing
    Set gobjCamposInvisiveis = Nothing
    Set gcolEdicaoTela = Nothing
    Set GL_objExcelSist = Nothing

End Sub

Public Property Get giSQLTipoOrdParam() As Integer
    giSQLTipoOrdParam = SistGlob_iSQLTipoOrdParam
End Property

Public Property Get giSQLTipoOrdParamDic() As Integer
    giSQLTipoOrdParamDic = SistGlob_iSQLTipoOrdParamDic
End Property

Public Property Let giSQLTipoOrdParam(ByVal vData As Integer)
    SistGlob_iSQLTipoOrdParam = vData
End Property

Public Property Let giSQLTipoOrdParamDic(ByVal vData As Integer)
    SistGlob_iSQLTipoOrdParamDic = vData
End Property

Sub Rotina_Erro_Carregar_Msg(sErroId As String, sTipoErro As String)

Dim lErro As Long, lComando As Long, Y As New AdmSQL

On Error GoTo Erro_Rotina_Erro_Carregar_Msg

    If GL_lConexaoDic Then
    
        lComando = Y.Comando_AbrirExt(GL_lConexaoDic)
        If lComando = 0 Then Error 1111
        
        sTipoErro = String(255, 0)
        lErro = Y.Comando_Executar(lComando, "SELECT Descricao FROM Erros WHERE Codigo = ?", sTipoErro, sErroId)
        If lErro <> AD_SQL_SUCESSO Then Error 1112
             
        lErro = Y.Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 1113
        
        Call Y.Comando_Fechar(lComando)
        
    Else
    
        sTipoErro = sErroId
        
    End If

    Exit Sub
     
Erro_Rotina_Erro_Carregar_Msg:

    If lComando Then Call Y.Comando_Fechar(lComando)
    
    sTipoErro = sErroId
     
    Exit Sub

End Sub

'Incluído em 02/07/2001 por Luiz Gustavo de Freitas Nogueira
Public Property Get GL_objExcelSist() As Object
    Set GL_objExcelSist = SistGlob_objExcel
End Property

'Incluído em 02/07/2001 por Luiz Gustavo de Freitas Nogueira
Public Property Set GL_objExcelSist(ByVal vData As Object)
    Set SistGlob_objExcel = vData
End Property

''Incluido em 08/01/2002 por Tulio
'
'*****Lembrar de declarar a constante
'*****MAILSEPARADOR quando for usar...
'
'Funcao Email_Enviar
'Envia um e-mail para destinatario(s) sTo

'Descricao dos parametros:
'
'   sTo -> Destinatario(s)
'   sSubject -> Assunto da mensagem
'   sTexto -> Corpo da mensagem
'   sAnexo -> Path do arquivo a ser anexado
'   sAliasAnexo -> Com qual nome o arquivo saira na mensagem

'Public Function Email_Enviar(sTo As String, Optional sSubject As String, Optional sTexto As String, Optional sAnexo As String, Optional sAliasAnexo As String) As Long
'
'Dim MAPISession1 As Object
'Dim MAPIMessages1 As Object
'Dim iPosToken As Integer
'Dim sToAux As String
'Dim sLetraCorrente As String
'
'On Error GoTo Erro_Email_Enviar
'
'    'Se Destinatario não foi preenchido
'    If Len(Trim(sTo)) = 0 Then gError 97101
'
'    Set MAPISession1 = CreateObject("MSMAPI.MAPISESSION")
'    Set MAPIMessages1 = CreateObject("MSMAPI.MAPIMESSAGES")
'
'    'Estabelece uma conexao
''------------------------------
'
'    MAPISession1.LogonUI = True
'
'    'Impedindo que os e-mails sejam baixados no inicio da conexao
'    MAPISession1.DownLoadMail = False
'
'    MAPISession1.SignOn
'    MAPISession1.NewSession = True
'
''------------------------------
'
'    'Apos estabelecida a conexao
'    'um identificador da sessao eh
'    'criado e armazenado na propriedade
'    'SessionID do controle MAPISession.
'    'Esse Identificador deve ser passado
'    'para o controle MAPIMessages
'
'    MAPIMessages1.SessionID = MAPISession1.SessionID
'
'    'Criando a mensagem
'    MAPIMessages1.Compose
'
'    '--- a seguir, as propriedades que soh
'    'podem ser alteradas em tempo de execucao
'
'    'Atribuindo o destino
'
'    For iPosToken = 1 To Len(sTo)
'        sLetraCorrente = Mid(sTo, iPosToken, 1)
'
'        If sLetraCorrente = MAILSEPARADOR Then
'            MAPIMessages1.RecipAddress = sToAux
'            MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
'            sToAux = ""
'            sLetraCorrente = ""
'        End If
'
'        sToAux = sToAux + sLetraCorrente
'    Next
'
'    MAPIMessages1.RecipAddress = sToAux
'
'    'Atribuindo o subject se foi passado
'    If Len(Trim(sSubject)) <> 0 Then
'        MAPIMessages1.MsgSubject = sSubject
'    End If
'
'    'Atribuindo o texto se foi passado
'    If Len(Trim(sTexto)) <> 0 Then
'        MAPIMessages1.MsgNoteText = sTexto
'    End If
'
'    'Se o anexo foi passado...
'    If Len(Trim(sAnexo)) <> 0 Then
'        'Definir tipo de dados do anexo
'        MAPIMessages1.AttachmentType = mapData
'
'        'Indicar o path
'        MAPIMessages1.AttachmentPathName = sAnexo
'
'        'Nome (Alias) do anexo na mensagem se foi passado
'        If Len(Trim(sAliasAnexo)) <> 0 Then
'            MAPIMessages1.AttachmentName = sAliasAnexo
'        End If
'    End If
'
'    'Envio Efetivo da mensagem
'    MAPIMessages1.Send False
'
'    'Fechando a conexao
'    MAPISession1.SignOff
'
'    MAPISession1.NewSession = False
'
'    Email_Enviar = SUCESSO
'
'    Exit Function
'
'Erro_Email_Enviar:
'
'    Email_Enviar = gErr
'
'    Select Case gErr
'
'        Case 97101
'            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142029)
'
'    End Select
'
'    MAPISession1.SignOff
'
'    MAPISession1.NewSession = False
'
'    Exit Function
'
'End Function

#If JONES_280115 = 1 Then
Private Sub Stack_GravaLog()

Dim lErro As Long
Dim iIndice As Integer
Dim sDiretorio As String
Dim sNomeArq As String
Dim sRegistro As String
Dim lRetorno As Long
Dim bArqAberto As Boolean
Dim iPos As Integer
Dim iPosAnt As Integer

On Error GoTo Erro_Stack_GravaLog

    bArqAberto = False

    sNomeArq = "StackLog_" & App.EXEName & UCase(gsNomeEmpresa) & "_" & UCase(gsUsuario) & gsLogonId & ".txt"
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "ArqLog", "C:\SGE\Programa\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    iPos = InStr(1, sDiretorio, "/")
    If iPos = 0 Then iPos = InStr(1, sDiretorio, "\")
    iPosAnt = iPos
    Do While iPos <> 0
        iPos = InStr(iPosAnt + 1, sDiretorio, "/")
        If iPos = 0 Then iPos = InStr(iPosAnt + 1, sDiretorio, "\")
        If iPos <> 0 Then iPosAnt = iPos
    Loop
    
    If iPosAnt <> 0 Then sDiretorio = left(sDiretorio, iPosAnt)

    If giNumCallStack = 0 Then
    
        Kill sDiretorio & sNomeArq
    
    Else
    
        Open sDiretorio & sNomeArq For Output As #1
        bArqAberto = True
        
        For iIndice = 1 To giNumCallStack
            Print #1, gsCallStack(iIndice)
        Next
        
        Close #1
        bArqAberto = False

    End If
    
    Exit Sub

Erro_Stack_GravaLog:

    Select Case gErr

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 999999)

    End Select

    If bArqAberto Then Close #1

    Exit Sub

End Sub

Private Sub CF_EntraNaStack(ByVal sNomeFuncao As String)

    giNumCallStack = giNumCallStack + 1
    gsCallStack(giNumCallStack) = sNomeFuncao
    Call Stack_GravaLog

End Sub

Private Sub CF_SaiDaStack(ByVal sNomeFuncao As String)
    
    gsCallStack(giNumCallStack) = ""
    giNumCallStack = giNumCallStack - 1
    Call Stack_GravaLog

End Sub
#End If

Function CF(sNomeFuncao As String, ParamArray vParams()) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_CF

#If JONES_280115 = 1 Then
    Call CF_EntraNaStack(sNomeFuncao)
#End If

    'alteracao feita por mario em 20/09/03 pois dava erro se algum codigo do ECF referenciava algum CF como em Grid_Trata_Tecla_Campo1
    'alterado por mario em 25/11/2014 pois o browser passou a ser chamado pelo ECF e ele usa CF
    If giLocalOperacao = LOCALOPERACAO_ECF Then
        If UBound(vParams) < 0 Then
            CF = CF_ECF(sNomeFuncao)
        ElseIf UBound(vParams) < 1 Then
            CF = CF_ECF(sNomeFuncao, vParams(0))
        ElseIf UBound(vParams) < 2 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1))
        ElseIf UBound(vParams) < 3 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2))
        ElseIf UBound(vParams) < 4 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3))
        ElseIf UBound(vParams) < 5 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
        ElseIf UBound(vParams) < 6 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5))
        ElseIf UBound(vParams) < 7 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6))
        ElseIf UBound(vParams) < 8 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7))
        ElseIf UBound(vParams) < 9 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8))
        ElseIf UBound(vParams) < 10 Then
            CF = CF_ECF(sNomeFuncao, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9))
        
        End If
        If lErro <> SUCESSO Then gError 214928
    Else

        lErro = Rotina_ObterObjetoFuncao(sNomeFuncao, X)
        If lErro <> SUCESSO Then gError 32276
    
    
        If UBound(vParams) < 0 Then
            CF = CallByName(X, sNomeFuncao, VbMethod)
        ElseIf UBound(vParams) < 1 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0))
        ElseIf UBound(vParams) < 2 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1))
        ElseIf UBound(vParams) < 3 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2))
        ElseIf UBound(vParams) < 4 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3))
        ElseIf UBound(vParams) < 5 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
        ElseIf UBound(vParams) < 6 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5))
        ElseIf UBound(vParams) < 7 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6))
        ElseIf UBound(vParams) < 8 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7))
        ElseIf UBound(vParams) < 9 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8))
        ElseIf UBound(vParams) < 10 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9))
        ElseIf UBound(vParams) < 11 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10))
        ElseIf UBound(vParams) < 12 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11))
        ElseIf UBound(vParams) < 13 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12))
        ElseIf UBound(vParams) < 14 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13))
        ElseIf UBound(vParams) < 15 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14))
        ElseIf UBound(vParams) < 16 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15))
        ElseIf UBound(vParams) < 17 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16))
        ElseIf UBound(vParams) < 18 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17))
        ElseIf UBound(vParams) < 19 Then
            CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17), vParams(18))
        Else
            gError 32277
        End If

    End If

#If JONES_280115 = 1 Then
    Call CF_SaiDaStack(sNomeFuncao)
#End If

    Exit Function

Erro_CF:

    CF = gErr

    Select Case gErr

        Case 13
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " está com tipo de parametro incorreto")
        
        Case 438
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " nao foi encontrada")
        
        Case 440
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, "chamada por CF de " & sNomeFuncao)
        
        Case 449
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " inclui parametro nao opcional")
        
        Case 32276
            lErro = Rotina_Erro(vbOKOnly, "Erro na via CF a " & sNomeFuncao, gErr)

        Case 32277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SUPERIOR_DE_PARAMETROS", gErr)

        Case 214928

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$ & " : " & sNomeFuncao, 142030)

    End Select

#If JONES_280115 = 1 Then
    Call CF_SaiDaStack(sNomeFuncao)
#End If

    Exit Function

End Function

Public Function Grid_Replica_Campo(objGridInt As AdmGrid, iTecla As Integer)
'Recebe a posição do campo atual, copia o conteúdo do campo acima e "cola" no campo atual
'objGridInt RECEBE(Input) as informações sobre o grid selecionado e sobre a tela desse grid
'iTecla RETORNA(Output) a tecla que será necessário simular a digitação

Dim objControleAtual As Object
Dim sValorCelulaAnterior As String
Dim sTipoControle As String
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer
Dim iIndice As Integer
Dim bAchou As Boolean
Dim bFocoNoGrid As Boolean
Dim sMascaraProduto As String
Dim dPercentual As Double

On Error GoTo Erro_Grid_Replica_Campo

     'Guarda o número da linha atual
    iLinhaAtual = objGridInt.objGrid.Row
    
    'Guarda o número da coluna atual
    iColunaAtual = objGridInt.objGrid.Col
    
    'Instancia objControleAtual como o controle selecionado no grid
    Set objControleAtual = objGridInt.objForm.Controls(objGridInt.colCampo(iColunaAtual))
    
    'Se o controle não estiver habilitado => sai da função, pois não faz sentido replicar um campo
    'não habilitado
    If objControleAtual.Enabled = False Then Exit Function
    
    'Se o número da linha atual menos o número de linhas fixas for maior que zero =>
    'significa que o conteúdo pode ser repetido, pois a linha ativa não é a primeira do grid
    If iLinhaAtual - objGridInt.objGrid.FixedRows > 0 Then
    
        '*** Descobre se o controle está no grid ***
        'Para cada coluna campo na coleção, ou seja, para cada coluna do grid
        For iIndice = 1 To objGridInt.colCampo.Count
            
            'Verica se o nome do controle ativo está na coleção, ou seja, verifica se
            'é um controle do grid
            If objGridInt.objForm.ActiveControl.Name = objGridInt.colCampo(iIndice) Then
                
                'Indica que achou o controle no grid
                bAchou = True
                
                'Sai do loop, pois o controle já foi encontrado
                Exit For
            End If
        
        Next
            
        'Se o controle ativo não é o próprio grid e não foi encontrado na coleção de campos do grid =>
        'significa que pode sair da função pois não faz sentido replicar o campo
        If (objGridInt.objForm.ActiveControl.Name <> objGridInt.objGrid.Name) And Not bAchou Then
            Exit Function
        ElseIf Not bAchou Then
            bFocoNoGrid = True
        End If
        '********************************************
        
        'Guarda em sValorCelulaAnterior o valor do campo imediatamente acima do controle atual
        sValorCelulaAnterior = objGridInt.objGrid.TextMatrix(iLinhaAtual - 1, iColunaAtual)
        
        'Guarda o tipo de controle
        sTipoControle = TypeName(objControleAtual)
        
        '*** Faz a cópia do controle anterior para o controle atual ***
        'Verifica o tipo do controle que está sendo replicado
        Select Case sTipoControle
        
            'Se for uma ComboBox
            Case CONTROLE_COMBOBOX
            
                'Limpa o item selecionado
                objControleAtual.ListIndex = -1
                
                'Se for uma combo estilo DropDown => Limpa o conteúdo exibido (Text)
                If objControleAtual.Style = vbComboDropdown Then objControleAtual.Text = ""
                
                'Se foi encontrado um valor na célula anterior, ou seja, se a combo acima da atual está preenchida
                If Len(sValorCelulaAnterior) > 0 Then
                
                    'Para cada item da combo
                    For iIndice = 0 To objControleAtual.ListCount - 1
                        
                        'Se o conteúdo do item for igual ao conteúdo da célula anterior
                        If objControleAtual.List(iIndice) = sValorCelulaAnterior Then
                            
                            'Seleciona o item na combo atual
                            objControleAtual.ListIndex = iIndice
                            
                            'Exibe no grid o conteúdo do campo atual
                            objGridInt.objGrid.Text = sValorCelulaAnterior
                            
                            'Sai do loop, pois o item já foi encontrado
                            Exit For
                        
                        End If
                    
                    Next
                    
                    'Se o índice utilizado no loop é maior que o número de itens da combo,
                    'significa que o conteúdo da célula anterior não foi encontrado na combo atual...
                    If iIndice > objControleAtual.ListCount - 1 Then
                        
                        'portanto => erro
                        gError GRID_CONTEUDO_INVALIDO_PARA_REPLICAR
                    
                    End If
                
                End If
            
            'Se o controle for uma CheckBox
            Case CONTROLE_CHECKBOX
        
                'Guarda na CheckBox atual o valor da CheckBox anterior
                objControleAtual.Value = sValorCelulaAnterior
                
                'Exibe no grid o valor da CheckBox anterior
                objGridInt.objGrid.Text = sValorCelulaAnterior
                
                'Chama a função que marca a checkbox
                Call Grid_Refresh_Checkbox(objGridInt)
            
            'Se o controle for uma MaskEdBox
            Case CONTROLE_MASKEDBOX
        
                'Se a máscara utilizada for referente à data
                If objControleAtual.Mask = "##/##/##" Then
                    
                    'Guarda no controle atual o valor do controle anterior no formato de data
                    objControleAtual.Text = Format(sValorCelulaAnterior, "dd/mm/yy")
                    
                    'Exibe no grid o valor do controle anterior
                    objGridInt.objGrid.Text = sValorCelulaAnterior
                
                'Se tiver formato numérico, ou formato do estoque
                ElseIf (objControleAtual.Format = "#,##0.00") Or (objControleAtual.Format = FORMATO_ESTOQUE) Then
                    
                    'Guarda no controle atual o valor do controle anterior no formato genérico para números
                    objControleAtual.Text = Format(sValorCelulaAnterior, "General Number")
                    
                    'Exibe no grid o valor do controle anterior
                    objGridInt.objGrid.Text = sValorCelulaAnterior
                
                'Se tiver formato percentual
                ElseIf objControleAtual.Format = "0%" Then
                    
                    'Retira o formato percentual do valor
                    dPercentual = CDbl(Format(sValorCelulaAnterior, "General Number")) * 100
                    
                    'Guarda no controle atual o valor do controle anterior sem o formato percentual
                    objControleAtual.Text = CStr(dPercentual)
                    
                    'Exibe no grid o valor do controle anterior (com o formato percentual)
                    objGridInt.objGrid.Text = sValorCelulaAnterior
                
                'Se for outro tipo de máscara ou formato
                Else
                    
                    'Guarda no controle o valor do campo anterior
                    objControleAtual.Text = sValorCelulaAnterior
                    
                    'Exibe no grid o valor do campo anterior
                    objGridInt.objGrid.Text = sValorCelulaAnterior
                End If
            
            'Se for uma OptionButton => não faz nada, pois conteúdo de options
            'não tem sentido replicar
            Case CONTROLE_OPTIONBUTTON
            
            
            'Se for outro tipo de controle
            Case Else
                
                'Guarda no controle o valor do campo anterior
                objControleAtual.Text = sValorCelulaAnterior
                
                'Exibe no grid o valor do campo anterior
                objGridInt.objGrid.Text = sValorCelulaAnterior
        
        End Select
        
        'Se o número da linha atual menos o número de linhas fixas for igual ao número de linhas existentes =>
        'significa que esta é a última linha do grid e é necessário incluir outra
        If iLinhaAtual - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objGridInt.objGrid.Col))) > 0 Then
            
            'Inclui mais uma linha no grid
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
        End If
        
        'Se o nome do controle ativo for diferente do nome do grid =>
        If Not bFocoNoGrid Then
            
            'significa que o usuário entrou no campo e é necessário simular o efeito da tecla enter
            'para executar o Saida_Celula e tornar o próximo campo ativo
            'iTecla = GRID_TECLAR_ENTER ??? verificar se pode retirar essa linha
            SendKeys TECLA_ENTER, True
        
        'Senão
        Else
            'Significa que o usuário não entrou no campo e é necessário simular o efeito da tecla esc
            'para executar o Saida_Celula e manter o "retângulo" no mesmo campo
            'iTecla = GRID_TECLAR_ESC ??? verificar se pode retirar essa linha
            SendKeys TECLA_ENTER, False
            SendKeys TECLA_ESC, True

        End If
    
    End If
    
    Grid_Replica_Campo = SUCESSO
    
    Exit Function

Erro_Grid_Replica_Campo:

    Grid_Replica_Campo = gErr
    
    Select Case gErr
    
        Case GRID_CONTEUDO_INVALIDO_PARA_REPLICAR
            'ERRO_CONTEUDO_CAMPO_INVALIDO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142031)
    
    End Select
    
    Exit Function
    
End Function

'Criado por Cyntia
Function Rotina_AvisoECF2(ByVal MsgBoxTipo As VbMsgBoxStyle, sErroId As String, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As VbMsgBoxResult
''   Rotina_Aviso = MsgBox("Confirma a operação?", MsgBoxTipo, "Rotina temporaria de aviso")
Dim sErro As String
Dim lErroAux As Integer
Dim X As Object, sTipoErro As String
Dim iErro As Integer

On Error GoTo Erro_Rotina_AvisoECF2
    
    Rotina_AvisoECF2 = 0
    sErro = String(1024, 0)
       
    If sErroId = "AVISO_DESEJA_SALVAR_ALTERACOES" Then sErroId = "Deseja salvar as alterações realizadas?"
    If sErroId = "AVISO_CONFIRMA_EXCLUSAO_LINHA_GRID" Then sErroId = "Confirma a exclusão da linha %s ?"
    
    If Not IsMissing(vParam10) Then
        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
        If lErroAux Then Exit Function
    Else
        If Not IsMissing(vParam9) Then
            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam8) Then
                lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam7) Then
                    lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam6) Then
                        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam5) Then
                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam4) Then
                                lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam3) Then
                                    lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam2) Then
                                        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam1) Then
                                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        Else
                                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, "", "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Set X = CreateObject("TelasAdm.ClassTelasAdm")
    
    If X Is Nothing Then Error 58450
    If X.objFormMsgAviso Is Nothing Then Error 58451
    
    'Passa o Erro e o Tipo dos Botões
    X.objFormMsgAviso.sErro = sErro
    X.objFormMsgAviso.MsgBoxTipo = MsgBoxTipo
    
    X.objFormMsgAviso.Show vbModal
    
    'Retorna o Resultado da escolha
    Rotina_AvisoECF2 = X.objFormMsgAviso.MsgBoxResultado
    
'    Rotina_AvisoECF = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")

    Exit Function
    
Erro_Rotina_AvisoECF2:
    
    Select Case Err
    
        Case 58450, 58451
            Call MsgBox("Erro na exibição de uma pergunta (ou aviso).", vbOKOnly, "SGE - Forprint")
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Error 59060 'para que quem chamou nao prossiga "normalmente" na operacao
    
    Exit Function
    
End Function

Function Rotina_ErroECF2(ByVal MsgBoxTipo As Integer, ByVal sErroId As String, ByVal lCodigo As Long, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As Long

Dim sErro As String, X As Object, Y As New AdmSQL
Dim iMsgBoxRet As Integer, lCodigoInt As Long, sTipoErro As String
Dim lErroAux As Integer, iTipoCliente As Integer

On Error GoTo Erro_Rotina_ErroECF2

    iTipoCliente = Sistema_ObterTipoCliente(GL_lSistema)
    If iTipoCliente = AD_SIST_BATCH Then
        lCodigoInt = 0
    Else
        lCodigoInt = lCodigo
    End If
    
    'para depurar o batchest2 como uma dll o trecho abaixo até os asteriscos deve estar comentado
    'se há alguma transacao aberta vou forcar o rollback
    If GL_lTransacao <> 0 And iTipoCliente <> AD_SIST_BATCH Then
        Call Y.Transacao_Rollback
    End If
'****
    
    GL_lUltimoErro = lCodigo

    Rotina_ErroECF2 = 0
    sErro = String(1024, 0)


    'para depurar o batchest como uma dll o comando abaixo deve estar descomentado
'    iTipoCliente = AD_SIST_BATCH
'****
    'verifiva o tipo do erro para pegar a mensagem correnspondente
    Select Case sErroId
        Case "ERRO_FORNECIDO_PELO_VB": sErroId = "Erro fornecido pelo VB: %s."
        Case "ERRO_VALOR_PORCENTAGEM3": sErroId = "O valor %s não está entre 0 e 99."
        Case "ERRO_VALOR_PORCENTAGEM": sErroId = "O valor %s não está entre 0 e 100."
        Case "ERRO_VALOR_INVALIDO": sErroId = "O valor digitado: %s é inválido. Exemplos de formatos válidos: 12560, 12.560, 12.560,35, 12560,35."
        Case "ERRO_FORMATO_DATA": sErroId = "Data Inválida."
        Case "ERRO_VALOR_NAO_POSITIVO": sErroId = "O valor digitado tem que ser positivo. Valor %s."
        Case "ERRO_CGC_NAO_NUMERICO": sErroId = "Cgc %s não é numérico."
        Case "ERRO_CGC_MENOR_OU_IGUAL_ZERO": sErroId = "Cgc %s é nulo ou negativo."
        Case "ERRO_CGC_NAO_INTEIRO": sErroId = "Cgc %s não é inteiro."
        Case "ERRO_CGC_INVALIDO": sErroId = "%s é Cgc inválido."
        Case "ERRO_CGC_MAIOR_QUE_CGCMAXIMO": sErroId = "Cgc %s é maior do que o Cgc máximo."
        Case "ERRO_CPF_NAO_NUMERICO": sErroId = "Cpf %s não é numérico."
        Case "ERRO_CPF_MENOR_OU_IGUAL_ZERO": sErroId = "Cpf %s é nulo ou negativo."
        Case "ERRO_CPF_NAO_INTEIRO": sErroId = "Cpf %s não é inteiro."
        Case "ERRO_CPF_MAIOR_QUE_CPFMAXIMO": sErroId = "Cpf %s é maior do que o Cpf máximo."
        Case "ERRO_CPF_INVALIDO": sErroId = "%s é Cpf inválido."
        Case "ERRO_VALOR_NAO_NUMERICO": sErroId = "O valor %s tem que ser numérico."
        Case "ERRO_ABERTURA_TRANSACAO": sErroId = "Não conseguiu abrir a transação."
        Case "ERRO_ARQUIVO_NAO_ENCONTRADO": sErroId = "O arquivo %s não foi encontrado."
        Case "ERRO_PREENCHIMENTO_ARQUIVO_CONFIG": sErroId = "O item %s da seção %s do arquivo de configuração %s não foi preenchido corretamente."
    End Select
    
    If Not IsMissing(vParam10) Then
        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
        If lErroAux Then Exit Function
    Else
        If Not IsMissing(vParam9) Then
            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam8) Then
                lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam7) Then
                    lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam6) Then
                        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam5) Then
                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam4) Then
                                lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam3) Then
                                    lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam2) Then
                                        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam1) Then
                                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        Else
                                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, "", "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "Empresa: " & gsNomeEmpresa & " Filial: " & gsNomeFilialEmpresa & " - " & Format(Now(), "General Date") & " - " & "Usuário: " & gsUsuario)
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "   ERRO: " & sTipoErro & " Local: " & CStr(lCodigo) & " Descrição: " & sErro)
'
    If iTipoCliente <> AD_SIST_BATCH Then
        
'        Set X = CreateObject("TelasAdm.ClassTelasAdm")
'
'        If X Is Nothing Then Error 58452
'        If X.objFormMsgErro Is Nothing Then Error 58453
'
'        X.objFormMsgErro.sErro = sErro
'        X.objFormMsgErro.lLocalErro = lCodigo
'        X.objFormMsgErro.sTipoErro = sErroId
'
'        MsgBox "Fase 23 Rotina_ERROECF2"
'
'        X.objFormMsgErro.Show vbModal
        
        MsgBox sErro
        
        'Rotina_Erro = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")
        
    End If
    
    Rotina_ErroECF2 = vbOK

    Exit Function
    
Erro_Rotina_ErroECF2:
    
    Select Case Err
    
        Case 58452, 58453
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Exit Function

End Function

Function CF_ECF(sNomeFuncao As String, ParamArray vParams()) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_CF_ECF

    lErro = Rotina_ObterObjetoFuncao_ECF(sNomeFuncao, X)
    If lErro <> SUCESSO Then gError 99896

    If UBound(vParams) < 0 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod)
    ElseIf UBound(vParams) < 1 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0))
    ElseIf UBound(vParams) < 2 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1))
    ElseIf UBound(vParams) < 3 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2))
    ElseIf UBound(vParams) < 4 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3))
    ElseIf UBound(vParams) < 5 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
    ElseIf UBound(vParams) < 6 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5))
    ElseIf UBound(vParams) < 7 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6))
    ElseIf UBound(vParams) < 8 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7))
    ElseIf UBound(vParams) < 9 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8))
    ElseIf UBound(vParams) < 10 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9))
    ElseIf UBound(vParams) < 11 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10))
    ElseIf UBound(vParams) < 12 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11))
    ElseIf UBound(vParams) < 13 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12))
    ElseIf UBound(vParams) < 14 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13))
    ElseIf UBound(vParams) < 15 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14))
    ElseIf UBound(vParams) < 16 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15))
    ElseIf UBound(vParams) < 17 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16))
    ElseIf UBound(vParams) < 18 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17))
    ElseIf UBound(vParams) < 19 Then
        CF_ECF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17), vParams(18))
    Else
        gError 99697
    End If

    Exit Function

Erro_CF_ECF:

    CF_ECF = gErr

    Select Case gErr

        Case 13
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " está com tipo de parametro incorreto")
        
        Case 438
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " nao foi encontrada")
        
        Case 449
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " inclui parametro nao opcional")
        
        Case 99696
            lErro = Rotina_Erro(vbOKOnly, "Erro na via CF a " & sNomeFuncao, gErr)

        Case 99697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SUPERIOR_DE_PARAMETROS", gErr)

        Case 440
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, "chamada por CF de " & sNomeFuncao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$ & " : " & sNomeFuncao, 142032)
            
    End Select

    Exit Function

End Function

Function Rotina_ObterObjetoFuncao_ECF(sCodigoRotina As String, X As Object) As Long
'Retorna um objeto que implemente uma funcao que tem o local de sua implementacao e a autorizacao de acesso identificados por sCodigoRotina

Dim sProjeto As String, sClasse As String, lErro As Long, sTemp As String

On Error GoTo Erro_Rotina_ObterObjetoFuncao_ECF

    If mvarcolRotinasECF Is Nothing Then
    
        lErro = Leitura_Arquivo_ECF
        If lErro <> SUCESSO Then gError 99900
    
    End If
    
    'leitura da col do local da função
    If mvarcolRotinasECF.Exists(sCodigoRotina) = False Then gError 99900
    
    mvarvBufferObterRot = mvarcolRotinasECF(sCodigoRotina)
    
    'pesquisa se o objeto foi criado anteriormente e inserido na colecao
    If mvarColObjFunc.Exists(mvarvBufferObterRot) Then
        
        Set X = mvarColObjFunc(mvarvBufferObterRot)
    
    Else
    
        Set X = CreateObject(mvarvBufferObterRot)
        
        'carrega
        Call mvarColObjFunc.Add(mvarvBufferObterRot, X)
        
    End If
    
    Rotina_ObterObjetoFuncao_ECF = SUCESSO
     
    Exit Function
    
Erro_Rotina_ObterObjetoFuncao_ECF:

    Rotina_ObterObjetoFuncao_ECF = gErr
     
    Select Case gErr
          
        Case 99900
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL1", gErr, sCodigoRotina)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142033)
     
    End Select
     
    Exit Function

End Function

Function Leitura_Arquivo_ECF() As Long

Dim sNomeRotina As String
Dim sLocal As String
Dim sRegistro As String
Dim iPos As Integer
Dim sArquivo As String
Dim sArquivo1 As String
Dim lTamanho As Long

On Error GoTo Erro_Leitura_Arquivo_ECF
    
    Set mvarcolRotinasECF = CreateObject("Scripting.Dictionary")
    
    lTamanho = 128
    sArquivo = String(lTamanho, 0)
    
    'verifica se o caixa possui ECF
    Call GetPrivateProfileString("DADOS", "Rotinas", -1, sArquivo, lTamanho, "CAIXACONFIG.ini")
    If sArquivo = String(lTamanho, 0) Then gError 126750
    
    sArquivo = StringZ(sArquivo)
    
    'se o diretorio for terminado por \  ===> retirar
    If right(sArquivo, 1) = "\" Then sArquivo = left(sArquivo, Len(sArquivo) - 1)
    
    sArquivo1 = Dir(sArquivo)
    
    If Len(sArquivo1) = 0 Then gError 126751
    
    'Abre o arquivo de retorno
    Open sArquivo For Input As #1
   
    'Até chegar ao fim do arquivo
    Do While Not EOF(1)
    
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Procura o sinal para separar o nome da rotina do seu local
        iPos = InStr(1, sRegistro, "=")
        
        sNomeRotina = left(sRegistro, iPos - 1)
        sLocal = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos - 1))
        
        mvarcolRotinasECF.Add sNomeRotina, sLocal
        
    Loop
    
    Close #1
    
    Exit Function
    
Erro_Leitura_Arquivo_ECF:

Leitura_Arquivo_ECF = gErr
     
    Close 1
    
    Select Case gErr
          
        Case 126750
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_ARQUIVO_CONFIG", gErr, "Rotinas", "DADOS", "CAIXACONFIG.ini")
          
        Case 126751
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO", gErr, sArquivo)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142034)
     
    End Select
     
    Exit Function

End Function

Public Function Sistema_Fechar_ECF()

    Set mvarcolRotinasECF = Nothing
    Set GL_objKeepAlive = Nothing
    Set mvarColObjFunc = Nothing
End Function

'Alteracao Daniel em 07/11/02
Public Function Controle_ObterNomeClasse(obj As Object) As String

Dim sNomeClass As String
Dim iPos As Integer
Static bMsg As Boolean

On Error GoTo Erro_Controle_ObterNomeClasse

    Controle_ObterNomeClasse = TypeName(obj)

'''    sNomeClass = String(MAX_NOME_CLASSE_CTL, 0)
'''
'''    If TypeOf obj Is Label Then
'''        sNomeClass = "Label"
'''        Exit Function
'''    ElseIf TypeOf obj Is Image Then
'''        sNomeClass = "Image"
'''        Exit Function
'''    Else
'''        Call GetClassName(obj.hWnd, sNomeClass, MAX_NOME_CLASSE_CTL)
'''    End If
'''
'''    iPos = InStr(1, sNomeClass, Chr$(0), 0)
'''    If iPos > 0 Then
'''        sNomeClass = Mid$(sNomeClass, 1, iPos - 1)
'''    End If
'''
'''    If InStr(1, sNomeClass, "CommandButton") > 0 Then
'''        sNomeClass = "CommandButton"
'''    ElseIf InStr(1, sNomeClass, "CheckBox") > 0 Then
'''            sNomeClass = "CheckBox"
'''    ElseIf InStr(1, sNomeClass, "Mask") > 0 Then
'''            sNomeClass = "MaskEdBox"
'''    ElseIf InStr(1, sNomeClass, "ComboBox") > 0 Then
'''            sNomeClass = "ComboBox"
'''    ElseIf InStr(1, sNomeClass, "TextBox") > 0 Then
'''            sNomeClass = "TextBox"
'''    ElseIf InStr(1, sNomeClass, "OptionButton") > 0 Then
'''            sNomeClass = "OptionButton"
'''    ElseIf InStr(1, sNomeClass, "Frame") > 0 Then
'''            sNomeClass = "Frame"
'''    ElseIf InStr(1, sNomeClass, "Form") > 0 Then
'''            sNomeClass = "Form"
'''    ElseIf InStr(1, sNomeClass, "TreeView") > 0 Then
'''            sNomeClass = "TreeView"
'''    ElseIf InStr(1, sNomeClass, "MSFlexGrid") > 0 Then
'''            sNomeClass = "MSFlexGrid"
'''    ElseIf InStr(1, sNomeClass, "PictureBox") > 0 Then
'''            sNomeClass = "PictureBox"
'''    ElseIf InStr(1, sNomeClass, "UpDown") > 0 Then
'''            sNomeClass = "UpDown"
'''    ElseIf InStr(1, sNomeClass, "ListBox") > 0 Then
'''            sNomeClass = "ListBox"
'''    ElseIf InStr(1, sNomeClass, "UserControl") > 0 Then
'''            sNomeClass = "UserControl"
'''    ElseIf InStr(1, sNomeClass, "TabStrip") > 0 Then
'''            sNomeClass = "TabStrip"
'''    ElseIf InStr(1, sNomeClass, "VScrollBar") > 0 Then
'''            sNomeClass = "VScrollBar"
'''    ElseIf InStr(1, sNomeClass, "HScrollBar") > 0 Then
'''            sNomeClass = "HScrollBar"
'''    Else
'''            Call WritePrivateProfileString("ControlesNaoTratados", sNomeClass, sNomeClass, "ADM100.INI")
'''
'''            'Daniel: Esse trecho de msg é temporário apenas para fins de testes internos.
'''                        'Favor retirar após testes do suporte.
'''            If bMsg = False Then
'''                Call MsgBox("Existem controles não tratados. Verifique o ADM100.ini", vbOKOnly, "Forprint")
'''                bMsg = True
'''            End If
'''
'''
'''    End If
'''
'''    Controle_ObterNomeClasse = sNomeClass
    
    Exit Function
    
Erro_Controle_ObterNomeClasse:

    Select Case Err
    
'''        Case 438
'''        'caso não tenha um hWnd associado ... retorna em branco, pois os controles antes tratados com typeof estão no if-else
        
        Case Else
            MsgBox (CStr(Err) & " : " & Error)
        
    End Select
    
    Controle_ObterNomeClasse = ""
    
End Function

Public Function NaturezaOp_Conv34(ByVal sNatOp As String, ByVal dtDataRef As Date) As String
'converte uma natureza de operacao de 3 para 4 digitos se a data de referencia for maior ou igual a DATA_INICIO_CFOP4
'se a natureza já vier com 4 digitos nao precisa fazer nada
'??? esta é uma funcao provisoria em 12/2002. depois deve ser eliminada

Dim sNatopConv As String

    sNatopConv = sNatOp
    
    If Len(sNatOp) = 3 And ((dtDataRef = DATA_NULA And gdtDataAtual >= DATA_INICIO_CFOP4) Or dtDataRef >= DATA_INICIO_CFOP4) Then
        
        Select Case sNatOp
        
            Case "111"
                sNatopConv = "1101"
            
            Case "112"
                sNatopConv = "1102"
            
            Case "113"
                sNatopConv = "1125"
            
            Case "131"
                sNatopConv = "1201"
            
            Case "132"
                sNatopConv = "1202"
            
            Case "142"
                sNatopConv = "1252"
            
            Case "143"
                sNatopConv = "1253"
            
            Case "144"
                sNatopConv = "1254"
            
            Case "162"
                sNatopConv = "1352"
            
            Case "163"
                sNatopConv = "1353"
            
            Case "193"
                sNatopConv = "1901"
            
            Case "194"
                sNatopConv = "1902"
            
            Case "197"
                sNatopConv = "1556"
            
            Case "199"
                sNatopConv = "1949"
                        
            Case "511"
                sNatopConv = "5101"
            
            Case "512"
                sNatopConv = "5102"
                        
            Case "513"
                sNatopConv = "5124"
                        
            Case "531"
                sNatopConv = "5201"
                        
            Case "532"
                sNatopConv = "5202"
                
            Case "561"
                sNatopConv = "5352"
            
            Case "562"
                sNatopConv = "5353"
            
            Case "571"
                sNatopConv = "5401"
            
            Case "591"
                sNatopConv = "5551"
                        
            Case "593"
                sNatopConv = "5901"
            
            Case "595"
                sNatopConv = "5556"
                        
            Case "599"
                sNatopConv = "5949"
                        
        End Select
        
    End If
    
    NaturezaOp_Conv34 = sNatopConv
    
End Function

Public Function Natop_ErroTamanho(ByVal dtDataRef As Date, ByVal sNatOp As String) As Boolean

Dim bErro As Boolean

    bErro = False
    
    If dtDataRef <> DATA_NULA And dtDataRef < DATA_INICIO_CFOP4 And Len(Trim(sNatOp)) <> 3 Then bErro = True
    
    If (dtDataRef = DATA_NULA Or dtDataRef >= DATA_INICIO_CFOP4) And Len(Trim(sNatOp)) <> 4 Then bErro = True
    
    Natop_ErroTamanho = bErro
    
End Function

Function CF2(ByVal objCT As Object, sNomeFuncao As String, ParamArray vParams()) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_CF2

#If JONES_280115 = 1 Then
    Call CF_EntraNaStack(sNomeFuncao)
#End If

    'alteracao feita por mario em 20/09/03 pois dava erro se algum codigo do ECF referenciava algum CF como em Grid_Trata_Tecla_Campo1
    If giLocalOperacao = LOCALOPERACAO_ECF Then
        CF2 = SUCESSO
    Else

        lErro = Rotina_ObterObjetoFuncao2(objCT, sNomeFuncao, X)
        If lErro <> SUCESSO Then Error 32276
    
        If UBound(vParams) < 0 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT)
        ElseIf UBound(vParams) < 1 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0))
        ElseIf UBound(vParams) < 2 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1))
        ElseIf UBound(vParams) < 3 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2))
        ElseIf UBound(vParams) < 4 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3))
        ElseIf UBound(vParams) < 5 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
        ElseIf UBound(vParams) < 6 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5))
        ElseIf UBound(vParams) < 7 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6))
        ElseIf UBound(vParams) < 8 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7))
        ElseIf UBound(vParams) < 9 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8))
        ElseIf UBound(vParams) < 10 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9))
        ElseIf UBound(vParams) < 11 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10))
        ElseIf UBound(vParams) < 12 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11))
        ElseIf UBound(vParams) < 13 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12))
        ElseIf UBound(vParams) < 14 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13))
        ElseIf UBound(vParams) < 15 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14))
        ElseIf UBound(vParams) < 16 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15))
        ElseIf UBound(vParams) < 17 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16))
        ElseIf UBound(vParams) < 18 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17))
        ElseIf UBound(vParams) < 19 Then
            CF2 = CallByName(X, sNomeFuncao, VbMethod, objCT, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), vParams(17), vParams(18))
        Else
            Error 32277
        End If

    End If

#If JONES_280115 = 1 Then
    Call CF_SaiDaStack(sNomeFuncao)
#End If

    Exit Function

Erro_CF2:

    CF2 = Err

    Select Case Err

        Case 13
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " está com tipo de parametro incorreto")
        
        Case 438
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " nao foi encontrada")
        
        Case 449
            MsgBox ("A implementacao da funcao " & sNomeFuncao & " inclui parametro nao opcional")
        
        Case 32276
            lErro = Rotina_Erro(vbOKOnly, "Erro na via CF a " & sNomeFuncao, Err)

        Case 32277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SUPERIOR_DE_PARAMETROS", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$ & " : " & sNomeFuncao, 142035)

    End Select

#If JONES_280115 = 1 Then
    Call CF_SaiDaStack(sNomeFuncao)
#End If

    Exit Function

End Function

Function Rotina_ObterObjetoFuncao2(ByVal objCT As Object, sCodigoRotina As String, X As Object) As Long
'Retorna um objeto que implemente uma funcao que tem o local de sua implementacao e a autorizacao de acesso identificados por sCodigoRotina

Dim sProjeto As String, sClasse As String, lErro As Long, sTemp As String, sNomeClasse As String, iPosPonto As Integer

On Error GoTo Erro_Rotina_ObterObjetoFuncao2

    If Not gbVPN Then
    
        lErro = Rotina_ObterFuncaoInt(GL_lDicDados, sCodigoRotina, mvarvBufferObterRot)
        If (lErro <> AD_BOOL_TRUE) Then Error 59258
        
    Else
    
        lErro = Rotina_ObterFuncaoVB(sCodigoRotina)
        If lErro <> SUCESSO Then Error 59258
    
    End If
    
    iPosPonto = InStr(mvarvBufferObterRot, ".")
    If iPosPonto <> 0 Then
        sNomeClasse = Mid(mvarvBufferObterRot, iPosPonto + 1)
    End If
    
    If UCase(TypeName(objCT)) = UCase(sNomeClasse) Then
    
        Set X = objCT
    
    Else
    
        'pesquisa se o objeto foi criado anteriormente e inserido na colecao
        If mvarColObjFunc.Exists(mvarvBufferObterRot) Then
            
            Set X = mvarColObjFunc(mvarvBufferObterRot)
        
        Else
        
            Set X = CreateObject(mvarvBufferObterRot)
    
            'se nao for p/batch, deixar carregado
'            If InStr(1, mvarvBufferObterRot, "Batch", vbTextCompare) = 0 Then Call mvarColObjFunc.Add(mvarvBufferObterRot, X)
            
        End If
    
    End If
    
    Rotina_ObterObjetoFuncao2 = SUCESSO
     
    Exit Function
    
Erro_Rotina_ObterObjetoFuncao2:

    Rotina_ObterObjetoFuncao2 = Err
     
    Select Case Err
          
        Case 59258
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL1", Err, sCodigoRotina)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error & " : " & sCodigoRotina, 142036)
     
    End Select
     
    Exit Function

End Function

Function Saida_Celula_Padrao(ByVal objGridInt As AdmGrid, objControle As Object, Optional ByVal bIncluiLinha As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If bIncluiLinha Then
        If Len(Trim(objControle.Text)) > 0 Then
            lErro = Adiciona_Linha_Grid(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    End If
    
    Saida_Celula_Padrao = SUCESSO
     
    Exit Function
    
Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208634)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
     
    End Select
     
    Exit Function

End Function

Public Function Adiciona_Linha_Grid(ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Adiciona_Linha_Grid
              
    'verifica se precisa preencher o grid com uma nova linha
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Adiciona_Linha_Grid = SUCESSO
        
    Exit Function

Erro_Adiciona_Linha_Grid:

    Adiciona_Linha_Grid = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208637)

    End Select

    Exit Function

End Function

Function Rotina_ObterObjetoTela_ECF(sCodigoTela As String) As Long

Dim sProjeto As String, sClasse As String, lErro As Long, sTemp As String

On Error GoTo Erro_Rotina_ObterObjetoTela_ECF

    If mvarcolRotinasECF Is Nothing Then
    
        lErro = Leitura_Arquivo_ECF
        If lErro <> SUCESSO Then gError 199477
    
    End If
    
    'leitura da col do local da função
    If mvarcolRotinasECF.Exists(sCodigoTela) = True Then
    
        sCodigoTela = mvarcolRotinasECF(sCodigoTela)
        
    End If

    Rotina_ObterObjetoTela_ECF = SUCESSO
     
    Exit Function
    
Erro_Rotina_ObterObjetoTela_ECF:

    Rotina_ObterObjetoTela_ECF = gErr
     
    Select Case gErr
          
        Case 199477

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199478)
     
    End Select
     
    Exit Function

End Function

Public Sub Grava_Log(asReg() As String)

Dim lErro As Long
Dim iIndice As Integer
Dim sDiretorio As String
Dim sNomeArq As String
Dim sRegistro As String
Dim lRetorno As Long
Dim bArqAberto As Boolean
Dim iPos As Integer
Dim iPosAnt As Integer

On Error GoTo Erro_Grava_Log

    sNomeArq = UCase(gsNomeEmpresa) & "_" & UCase(gsUsuario) & "_" & Format(Now, "YYYYMMDDHHMMSS") & ".txt"
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "ArqLog", "C:\SGE\Programa\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    iPos = InStr(1, sDiretorio, "/")
    If iPos = 0 Then iPos = InStr(1, sDiretorio, "\")
    iPosAnt = iPos
    Do While iPos <> 0
        iPos = InStr(iPosAnt + 1, sDiretorio, "/")
        If iPos = 0 Then iPos = InStr(iPosAnt + 1, sDiretorio, "\")
        If iPos <> 0 Then iPosAnt = iPos
    Loop
    
    If iPosAnt <> 0 Then sDiretorio = left(sDiretorio, iPosAnt)

    Open sDiretorio & sNomeArq For Output As #1
    bArqAberto = True
    
    For iIndice = LBound(asReg) To UBound(asReg)
        Print #1, asReg(iIndice)
    Next
    
    Close #1
    bArqAberto = False

    Exit Sub

Erro_Grava_Log:

    Select Case gErr

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 999999)

    End Select

    If bArqAberto Then Close #1

    Exit Sub

End Sub

Public Function Arredonda_Moeda(dValor As Double, Optional ByVal iNumDigitos As Integer = 2) As Double

    If dValor >= 0 Then
        Arredonda_Moeda = Round(dValor + 0.0000000001, iNumDigitos)
    Else
        Arredonda_Moeda = Round(dValor - 0.0000000001, iNumDigitos)
    End If

End Function

Public Sub Grava_Log2(ByVal sTexto As String)

Dim lErro As Long
Dim bArqAberto As Boolean

On Error GoTo Erro_Grava_Log2

    Open "C:\SGE\LOG_DEBUG" & Format(Now, "YYYYMMDD") & ".txt" For Append As #171
    bArqAberto = True

    Print #171, Format(Now, "YYYYMMDDHHMMSS") & " - " & sTexto

    Close #171
    bArqAberto = False

    Exit Sub

Erro_Grava_Log2:

    Select Case gErr

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 999999)

    End Select

    If bArqAberto Then Close #171

    Exit Sub

End Sub

