VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchFPTEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Function Rotina_FPT_Envio(ByVal objLojaECF As ClassLojaECF) As Long
'rotina de envio dos arquivos gerados pelo loja para o FPT

Dim sRetorno As String
Dim lTamanho As Long
Dim sNomeArq1 As String
Dim sNomeArqErro As String
Dim sNomeArqEnv As String
Dim sDirEnv As String
Dim sArquivo As String

On Error GoTo Erro_Rotina_FPT_Envio

    If Len(objLojaECF.sFTPURL) > 0 Then

        lTamanho = 255
        sRetorno = String(lTamanho, 0)
        
        'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
        Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosECF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
        'Retira os espaços no final da string
        sRetorno = StringZ(sRetorno)
        
        'Se não encontrou
        If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127100
        
        sNomeArq1 = sRetorno & "\" & "*.ccc"
        '*************************************************************
        
        sNomeArqErro = sNomeArq1
        
        'retorna o primeiro arquivo com a extensao .ccc encontrado
        sArquivo = Dir(sNomeArq1)
        
        'verifica se o diretorio para onde vao ser copiados os arquivos existe
        'se nao existir, cria.
        sDirEnv = sRetorno & "\Env"
        
        If Not ((GetAttr(sDirEnv) And vbDirectory) = vbDirectory) Then
            MkDir sDirEnv
        End If
        
        Do While Len(sArquivo) > 0
            
            'comandos FPT
            
            'o arquivo vai ser transferido para este diretorio
            sNomeArqEnv = sDirEnv & "\" & sArquivo
            
            'eliminar o arquivo no diretorio "enviado" se existir
            Kill sNomeArqEnv
            
            'transferir o arquivo para sDirEnv
            Name sNomeArq1 As sNomeArqEnv
            
            'pegar o proximo arquivo
            sArquivo = Dir
        
        Loop

    End If

    Rotina_FPT_Envio = SUCESSO

    Exit Function

Erro_Rotina_FPT_Envio:

    Rotina_FPT_Envio = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144656)

    End Select

    Exit Function

End Function

