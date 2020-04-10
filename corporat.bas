Attribute VB_Name = "PrincCorporat"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const SUCESSO = 0
Public Const NOME_ARQUIVO_ADM = "ADM100.INI"
Public Const DLL_A_REGISTRAR = "Não há DLLs a Registrar"  'Inserir nome da DLL que deve ser registrada

Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SW_HIDE = 0

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
 
Public iParaTudo As Integer
Public colOutrosArquivos As New Collection

Sub Main()

Dim sAtualiza As String
Dim lErro As Long

On Error GoTo Erro_Principal

   iParaTudo = False

   sAtualiza = LeArqINI("Geral", "AutoAtualiza", NOME_ARQUIVO_ADM)
   If Len(sAtualiza) = 0 Then Error 5001

   If sAtualiza = "1" Then
   
        colOutrosArquivos.Add "DllInscE32.dll"
        colOutrosArquivos.Add "dicprincipal2.exe"
        colOutrosArquivos.Add "adrellib.dll"
        colOutrosArquivos.Add "RUNEXE4.dll"
        colOutrosArquivos.Add "D4GERAL.dll"
        colOutrosArquivos.Add "FWRUN40I.dll"
        colOutrosArquivos.Add "FWRUN40.dll"
        colOutrosArquivos.Add "FWRELRUN.dll"
        colOutrosArquivos.Add "FWODBC.dll"
        colOutrosArquivos.Add "FWLIB4.dll"
        colOutrosArquivos.Add "FORPWC4.dll"
        colOutrosArquivos.Add "F4YACCLK.dll"
        colOutrosArquivos.Add "F4EXTFPW.dll"
        colOutrosArquivos.Add "cnab_409.dll"
        colOutrosArquivos.Add "F4ATLOAD.dll"
        colOutrosArquivos.Add "cnab_341.dll"
        colOutrosArquivos.Add "cnab_237.dll"
        colOutrosArquivos.Add "cnab_001.dll"
        colOutrosArquivos.Add "DIFPW4.dll"
        colOutrosArquivos.Add "BTRVDLL4.dll"
        colOutrosArquivos.Add "D4ODBCU.dll"
        colOutrosArquivos.Add "D4ODBC.dll"
        colOutrosArquivos.Add "D4MSCOB.dll"
        colOutrosArquivos.Add "d4mbcob.dll"
        colOutrosArquivos.Add "D4FORP10.dll"
        colOutrosArquivos.Add "D4DBASE.dll"
        colOutrosArquivos.Add "D4CLANG.dll"
        colOutrosArquivos.Add "Adsqlmn.dll"
        colOutrosArquivos.Add "Adsqldrv.dll"
        colOutrosArquivos.Add "D4BASIC.dll"
        colOutrosArquivos.Add "D4ACUM.dll"
        colOutrosArquivos.Add "admaqexp.dll"
        colOutrosArquivos.Add "adinstal.dll"
        colOutrosArquivos.Add "adhelp02.dll"
        colOutrosArquivos.Add "adhelp01.dll"
        colOutrosArquivos.Add "adcusr.dll"
        colOutrosArquivos.Add "adcrtl.dll"
        colOutrosArquivos.Add "adcnab.dll"
        colOutrosArquivos.Add "cnab200.dll"
        colOutrosArquivos.Add "checkboxgrayed.bmp"
        colOutrosArquivos.Add "regsgef.bat"
        colOutrosArquivos.Add "RegCorporator.exe"
        colOutrosArquivos.Add "RegCorporator.bat"
        colOutrosArquivos.Add "ean3i.dll"
        colOutrosArquivos.Add "testeassinatura2.XmlSerializers.dll"
        colOutrosArquivos.Add "testeassinatura2.application"
        colOutrosArquivos.Add "testeassinatura2.exe.manifest"
        colOutrosArquivos.Add "testeassinatura2.exe"
        colOutrosArquivos.Add "testeassinatura2.pdb"
        colOutrosArquivos.Add "testeassinatura2.xml"
        colOutrosArquivos.Add "testeassinatura2.exe.config"
        
        colOutrosArquivos.Add "envEventoCancNFe_v1.00.xsd"
        colOutrosArquivos.Add "leiauteEventoCancNFe_v1.00.xsd"
        colOutrosArquivos.Add "cancNFe_v2.00.xsd"
        colOutrosArquivos.Add "cancNFe_v2.00n.xsd"
        colOutrosArquivos.Add "consCad_v2.00.xsd"
        colOutrosArquivos.Add "consReciNFe_v2.00.xsd"
        colOutrosArquivos.Add "consReciNFe_v2.00n.xsd"
        colOutrosArquivos.Add "consSitNFe_v2.00.xsd"
        colOutrosArquivos.Add "consSitNFe_v2.00n.xsd"
        colOutrosArquivos.Add "consStatServ_v2.00.xsd"
        colOutrosArquivos.Add "enviNFe_v2.00.xsd"
        colOutrosArquivos.Add "enviNFe_v2.00n.xsd"
        colOutrosArquivos.Add "inutNFe_v2.00.xsd"
        colOutrosArquivos.Add "inutNFe_v2.00n.xsd"
        colOutrosArquivos.Add "leiauteCancNFe_v2.00.xsd"
        colOutrosArquivos.Add "leiauteCancNFe_v2.00n.xsd"
        colOutrosArquivos.Add "leiauteConsSitNFe_v2.00.xsd"
        colOutrosArquivos.Add "leiauteConsSitNFe_v2.00n.xsd"
        colOutrosArquivos.Add "leiauteConsStatServ_v2.00.xsd"
        colOutrosArquivos.Add "leiauteConsultaCadastro_v2.00.xsd"
        colOutrosArquivos.Add "leiauteInutNFe_v2.00.xsd"
        colOutrosArquivos.Add "leiauteInutNFe_v2.00n.xsd"
        colOutrosArquivos.Add "leiauteNFe_v2.00.xsd"
        colOutrosArquivos.Add "leiauteNFe_v2.00n.xsd"
        colOutrosArquivos.Add "nfe_v2.00.xsd"
        colOutrosArquivos.Add "nfe_v2.00n.xsd"
        colOutrosArquivos.Add "procCancNFe_v2.00.xsd"
        colOutrosArquivos.Add "procCancNFe_v2.00n.xsd"
        colOutrosArquivos.Add "procInutNFe_v2.00.xsd"
        colOutrosArquivos.Add "procNFe_v2.00.xsd"
        colOutrosArquivos.Add "procNFe_v2.00n.xsd"
        colOutrosArquivos.Add "retCancNFe_v2.00.xsd"
        colOutrosArquivos.Add "retCancNFe_v2.00n.xsd"
        colOutrosArquivos.Add "retConsCad_v2.00.xsd"
        colOutrosArquivos.Add "retConsReciNFe_v2.00.xsd"
        colOutrosArquivos.Add "retConsReciNFe_v2.00n.xsd"
        colOutrosArquivos.Add "retConsSitNFe_v2.00.xsd"
        colOutrosArquivos.Add "retConsSitNFe_v2.00n.xsd"
        colOutrosArquivos.Add "retConsStatServ_v2.00.xsd"
        colOutrosArquivos.Add "retEnviNFe_v2.00.xsd"
        colOutrosArquivos.Add "retEnviNFe_v2.00n.xsd"
        colOutrosArquivos.Add "retInutNFe_v2.00.xsd"
        colOutrosArquivos.Add "retInutNFe_v2.00n.xsd"
        colOutrosArquivos.Add "testeassinatura4.application"
        colOutrosArquivos.Add "testeassinatura4.exe"
        colOutrosArquivos.Add "testeassinatura4.exe.config"
        colOutrosArquivos.Add "testeassinatura4.exe.manifest"
        colOutrosArquivos.Add "testeassinatura4.pdb"
        colOutrosArquivos.Add "testeassinatura4.xml"
        colOutrosArquivos.Add "testeassinatura4.XmlSerializers.dll"
        colOutrosArquivos.Add "tiposBasico_v1.03.xsd"
        colOutrosArquivos.Add "xmldsig-core-schema_v1.01.xsd"

        colOutrosArquivos.Add "nfse1.xsd"

        colOutrosArquivos.Add "nfseabrasf.exe"
        colOutrosArquivos.Add "nfseabrasf.exe.config"
        colOutrosArquivos.Add "nfseabrasf.xml"
        colOutrosArquivos.Add "nfseabrasf.XmlSerializers.dll"

        colOutrosArquivos.Add "nfserj.exe"
        colOutrosArquivos.Add "nfserj.exe.config"
        colOutrosArquivos.Add "nfserj.xml"
        colOutrosArquivos.Add "nfserj.XmlSerializers.dll"

        colOutrosArquivos.Add "nfsetatui.exe"
        colOutrosArquivos.Add "nfsetatui.exe.config"
        colOutrosArquivos.Add "nfsetatui.xml"
        colOutrosArquivos.Add "nfsetatui.XmlSerializers.dll"

        colOutrosArquivos.Add "nfsetatui2.exe"
        'colOutrosArquivos.Add "nfsetatui.exe.config"
        colOutrosArquivos.Add "nfsetatui2.xml"
        colOutrosArquivos.Add "nfsetatui2.XmlSerializers.dll"

        colOutrosArquivos.Add "ErrosBatch.exe"
        colOutrosArquivos.Add "azip32.dll"
        
        colOutrosArquivos.Add "Ionic.Zip.dll"
        
        colOutrosArquivos.Add "AtualizaCorporator.application"
        colOutrosArquivos.Add "AtualizaCorporator.exe"
        colOutrosArquivos.Add "AtualizaCorporator.exe.config"
        colOutrosArquivos.Add "AtualizaCorporator.exe.manifest"
        colOutrosArquivos.Add "AtualizaCorporator.pdb"
        colOutrosArquivos.Add "AtualizaCorporator.vshost.exe"
        colOutrosArquivos.Add "AtualizaCorporator.vshost.exe.config"
        colOutrosArquivos.Add "AtualizaCorporator.xml"

        colOutrosArquivos.Add "SGECorporator.application"
        colOutrosArquivos.Add "SGECorporator.exe"
        colOutrosArquivos.Add "SGECorporator.exe.config"
        colOutrosArquivos.Add "SGECorporator.exe.manifest"
        colOutrosArquivos.Add "SGECorporator.pdb"
        colOutrosArquivos.Add "SGECorporator.vshost.exe"
        colOutrosArquivos.Add "SGECorporator.xml"

        colOutrosArquivos.Add "SGEUpdate.application"
        colOutrosArquivos.Add "SGEUpdate.exe"
        colOutrosArquivos.Add "SGEUpdate.exe.config"
        colOutrosArquivos.Add "SGEUpdate.exe.manifest"
        colOutrosArquivos.Add "SGEUpdate.pdb"
        colOutrosArquivos.Add "SGEUpdate.vshost.exe"
        colOutrosArquivos.Add "SGEUpdate.vshost.exe.config"
        colOutrosArquivos.Add "SGEUpdate.xml"
        
        colOutrosArquivos.Add "Import_Xml.application"
        colOutrosArquivos.Add "Import_Xml.exe"
        colOutrosArquivos.Add "SGEImportXml.exe"
        colOutrosArquivos.Add "Import_Xml.exe.config"
        colOutrosArquivos.Add "Import_Xml.exe.manifest"
        colOutrosArquivos.Add "Import_Xml.pdb"
        colOutrosArquivos.Add "Import_Xml.xml"
        
        colOutrosArquivos.Add "Corporator_Zip.application"
        colOutrosArquivos.Add "Corporator_Zip.exe"
        colOutrosArquivos.Add "Corporator_Zip.exe.config"
        colOutrosArquivos.Add "Corporator_Zip.exe.manifest"
        colOutrosArquivos.Add "Corporator_Zip.pdb"
        colOutrosArquivos.Add "Corporator_Zip.xml"
        
        colOutrosArquivos.Add "CCe.xsd"
        colOutrosArquivos.Add "CCe_v1.00.xsd"
        colOutrosArquivos.Add "envCCe.xsd"
        colOutrosArquivos.Add "envCCe_v1.00.xsd"
        colOutrosArquivos.Add "leiau.xsd"
        colOutrosArquivos.Add "leiauteCCe_v1.00.xsd"
        colOutrosArquivos.Add "procCCe.xsd"
        colOutrosArquivos.Add "procCCeNFe_v1.00.xsd"
        colOutrosArquivos.Add "retEnvCCe.xsd"
        colOutrosArquivos.Add "retEnvCCe_v1.00.xsd"
        
        colOutrosArquivos.Add "consSitNFe_v2.01.xsd"
        colOutrosArquivos.Add "leiauteConsSitNFe_v2.01.xsd"
        colOutrosArquivos.Add "retConsSitNFe_v2.01.xsd"
        colOutrosArquivos.Add "tiposBasico_v1.03.xsd"
        colOutrosArquivos.Add "xmldsig-core-schema_v1.01.xsd"
        
        colOutrosArquivos.Add "suporteonline.htm"
        
        colOutrosArquivos.Add "sgenfebd.exe"
        colOutrosArquivos.Add "sgenfebd.exe.config"
        colOutrosArquivos.Add "sgenfebd.xml"
        
        colOutrosArquivos.Add "sgenfebd4.exe"
        colOutrosArquivos.Add "sgenfebd4.exe.config"
        colOutrosArquivos.Add "sgenfebd4.xml"
        
        colOutrosArquivos.Add "sgenfse.exe"
        colOutrosArquivos.Add "sgenfse.exe.config"
        colOutrosArquivos.Add "sgenfse.xml"
        
        
        colOutrosArquivos.Add "Interop.admlib.dll"
        colOutrosArquivos.Add "Interop.GlobaisAdm.dll"
        colOutrosArquivos.Add "Interop.GlobaisContab.dll"
        colOutrosArquivos.Add "Interop.GlobaisCRFAT.dll"
        colOutrosArquivos.Add "Interop.GlobaisFAT.dll"
        colOutrosArquivos.Add "Interop.GlobaisLoja.dll"
        colOutrosArquivos.Add "Interop.GlobaisMAT.dll"
        colOutrosArquivos.Add "Interop.GlobaisPV.dll"
        colOutrosArquivos.Add "Interop.GlobaisTRB.dll"
        colOutrosArquivos.Add "Interop.VBA.dll"
        
        colOutrosArquivos.Add "sgenfe.dll"
        colOutrosArquivos.Add "sgenfe.dll.config"
        colOutrosArquivos.Add "sgenfe.pdb"
        colOutrosArquivos.Add "sgenfe.tlb"
        colOutrosArquivos.Add "sgenfe.xml"
        colOutrosArquivos.Add "sgenfe.XmlSerializers.dll"
        
        colOutrosArquivos.Add "sgenfe4.dll"
        colOutrosArquivos.Add "sgenfe4.dll.config"
        colOutrosArquivos.Add "sgenfe4.pdb"
        colOutrosArquivos.Add "sgenfe4.tlb"
        colOutrosArquivos.Add "sgenfe4.xml"
        colOutrosArquivos.Add "sgenfe4.XmlSerializers.dll"
        
        colOutrosArquivos.Add "SGESAT.dll"
        colOutrosArquivos.Add "SGESAT.pdb"
        colOutrosArquivos.Add "SGESAT.tlb"
        colOutrosArquivos.Add "SGESAT.xml"
        
        colOutrosArquivos.Add "SGEUtil.dll"
        colOutrosArquivos.Add "SGEUtil.pdb"
        colOutrosArquivos.Add "SGEUtil.tlb"
        colOutrosArquivos.Add "SGEUtil.xml"
        
        lErro = Verifica_Atualizacoes(colOutrosArquivos)
        If lErro <> SUCESSO Then Error 5003

   End If
   
    If bExisteFrmWrk Then
        Shell "SGECorporator.exe", vbMaximized
    Else
        Shell "SGEPrinc2.exe", vbMaximized
    End If
   
   End

Erro_Principal:

   Beep
   
   Select Case Err.Number
   
      Case Is = 5001
         MsgBox "Não há informação sobre o atualização automática no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
      Case Is = 5003
      
      Case Else
         MsgBox "Erro do VB! (" & Err.Number & " - " & Err.Description & ").", vbCritical, "Atenção!"
      
   End Select
   
   End

End Sub

Public Function bExisteFrmWrk() As Boolean

Dim sDir As String
Dim sUsaFrmWrk As String

On Error GoTo Erro_bExisteFrmWrk

    'Incluído para evitar testes em instalações de Framwork problemáticas
    sUsaFrmWrk = LeArqINI("Geral", "UsaFrmWrk", NOME_ARQUIVO_ADM)

    If sUsaFrmWrk = "0" Then
        bExisteFrmWrk = False
    Else
        sDir = GetWinDir & "Microsoft.NET\Framework"
        
        If Len(Trim(Dir(sDir, vbDirectory))) = 0 Then
            bExisteFrmWrk = False
        Else
            bExisteFrmWrk = True
        End If
    End If

    Exit Function
    
Erro_bExisteFrmWrk:

    bExisteFrmWrk = False
    
    Exit Function

End Function

Private Function GetWinDir() As String
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(255, 0)
  lngResult = GetWindowsDirectory(strFolder, 255)
  If lngResult <> 0 Then
    If Right(Left(strFolder, lngResult), 1) = "\" Then
      GetWinDir = Left(strFolder, lngResult)
    Else
      GetWinDir = Left(strFolder, lngResult) & "\"
    End If
  Else
    GetWinDir = ""
  End If
End Function

Public Function LeArqINI(sSecao As String, sItem As String, sArqIni As String) As String
'Lê valor string no arquivo ADM100.INI
Dim sConteudo As String * 128
Dim lRetorno As Long
   
On Error GoTo Erro_LeArqINI
   
   lRetorno = GetPrivateProfileString(sSecao, sItem, "", sConteudo, 128, sArqIni)
   
   LeArqINI = Left(sConteudo, lRetorno)

   Exit Function
   
Erro_LeArqINI:

   LeArqINI = ""

   Exit Function

End Function

Public Function Extrai_Diretorio(sNomeArq As String) As String
'Extrai o nome do diretório do nome do arquivo
Dim iUltimaPos As Integer
Dim iBarraPos As Integer

On Error GoTo Erro_Extrai_Diretorio

   iUltimaPos = 2                     'depois do c:
   
   iBarraPos = InStr(sNomeArq, "\")
   
   Do While iBarraPos > 0
      
      iUltimaPos = iBarraPos
      iBarraPos = InStr(iBarraPos + 1, sNomeArq, "\")
   
   Loop
   
   Extrai_Diretorio = Mid(sNomeArq, 1, iUltimaPos)
   
   Exit Function

Erro_Extrai_Diretorio:

   Extrai_Diretorio = ""
   
   Exit Function
   
End Function

Public Function Verifica_Atualizacoes(colOutrosArquivos As Collection) As Long

Dim lErro As Long
Dim sArqLote As String
Dim sDirDestino As String
Dim sDirOrigem As String
Dim lArqNum As Long
Dim sLinha As String
Dim sNomeArq As String
Dim iEhEXE As Integer
Dim dDataArqOrigem As Date
Dim dDataArqDestino As Date
Dim iCopiaArq As Integer
Dim iCarregaForm As Integer
Dim iIndice As Integer
Dim sMsgErro As String, sMsgErroAux As String, iLinha As Integer, iErrosPnt As Integer

On Error GoTo Erro_Verifica_Atualizacoes

    sMsgErro = "Função: Verifica_Atualizacoes"
iErrosPnt = 1
   
    sArqLote = LeArqINI("Geral", "RegFile", NOME_ARQUIVO_ADM)
iErrosPnt = 2
    If Len(sArqLote) = 0 Then Error 5005
iErrosPnt = 3
    
    sMsgErro = sMsgErro & " ArqReg: " & sArqLote
iErrosPnt = 4
          
    sDirDestino = App.Path
iErrosPnt = 5
    If Right(sDirDestino, 1) <> "\" Then
iErrosPnt = 6
       sDirDestino = sDirDestino & "\"
iErrosPnt = 7
    End If
iErrosPnt = 8
    sMsgErro = sMsgErro & " Destino: " & sDirDestino
iErrosPnt = 9
    
    sDirOrigem = LeArqINI("Geral", "DirProgram", NOME_ARQUIVO_ADM)
iErrosPnt = 10
    If Len(sDirOrigem) = 0 Then Error 5007
iErrosPnt = 11
    
    If Right(sDirOrigem, 1) <> "\" Then
iErrosPnt = 12
       sDirOrigem = sDirOrigem & "\"
iErrosPnt = 13
    End If
iErrosPnt = 14
    sMsgErro = sMsgErro & " Origem: " & sDirOrigem
iErrosPnt = 15

    lErro = Teste_Acesso_Full_Pasta(sDirDestino)
    If lErro <> SUCESSO Then Error 10002
   
    iCarregaForm = False
iErrosPnt = 16
   
    lArqNum = FreeFile
iErrosPnt = 17
    
    Open sArqLote For Input As lArqNum
iErrosPnt = 18
 
    Do While Not EOF(lArqNum)
iErrosPnt = 19
    
Le_Outro:
iErrosPnt = 20

        iLinha = iLinha + 1
iErrosPnt = 21
        
        sMsgErroAux = "Tipo: Reg Linha: " & CStr(iLinha)
iErrosPnt = 22
   
        If EOF(lArqNum) Then Exit Do
iErrosPnt = 23
        
        Line Input #1, sLinha
iErrosPnt = 24
      
        sMsgErroAux = sMsgErroAux & " Conteudo: " & sLinha
iErrosPnt = 25
      
        iEhEXE = False
iErrosPnt = 26
      
        If UCase(Left(sLinha, 12)) = "REGSVR32 /S " Then
iErrosPnt = 27
          
            sNomeArq = Mid(sLinha, 13, (Len(sLinha) - 12))
iErrosPnt = 28
       
        ElseIf UCase(Right(sLinha, 15)) = ".EXE /REGSERVER" Then
iErrosPnt = 29
          
            sNomeArq = Mid(sLinha, 1, (Len(sLinha) - 11))
iErrosPnt = 30
            iEhEXE = True
iErrosPnt = 31
       
        Else
iErrosPnt = 32
        
           GoTo Le_Outro
iErrosPnt = 33
       
        End If
iErrosPnt = 34
          
        sMsgErroAux = sMsgErroAux & " NomeArq: " & sNomeArq
iErrosPnt = 35
          
        If Len(Dir(sDirOrigem & sNomeArq)) = 0 Then Error 5008
iErrosPnt = 36


        lErro = Testa_Arquivo(sDirOrigem & sNomeArq)
        If lErro <> SUCESSO Then Error 9999

       
        sMsgErroAux = sMsgErroAux & " NomeArq: " & sNomeArq
iErrosPnt = 37
       
        dDataArqOrigem = FileDateTime(sDirOrigem & sNomeArq)
iErrosPnt = 38
        
        sMsgErroAux = sMsgErroAux & " Date Origem: " & Format(dDataArqOrigem, "dd/mm/yyyy")
iErrosPnt = 39

        
        If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
iErrosPnt = 40
           
           iCarregaForm = True
iErrosPnt = 41
           Exit Do
iErrosPnt = 42
        
        Else
iErrosPnt = 43
        
           If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
iErrosPnt = 44
              iCarregaForm = True
iErrosPnt = 45
              Exit Do
iErrosPnt = 46
           End If
iErrosPnt = 47

            lErro = Testa_Arquivo(sDirDestino & sNomeArq)
            If lErro <> SUCESSO Then Error 9999
           
           dDataArqDestino = FileDateTime(sDirDestino & sNomeArq)
iErrosPnt = 48
        
            sMsgErroAux = sMsgErroAux & " Date Origem: " & Format(dDataArqDestino, "dd/mm/yyyy")
iErrosPnt = 49
        
           If dDataArqOrigem > dDataArqDestino Then
iErrosPnt = 50
              iCarregaForm = True
iErrosPnt = 51
              Exit Do
iErrosPnt = 52
           End If
iErrosPnt = 53
        
        End If
iErrosPnt = 54
        
    Loop
iErrosPnt = 55
    
    Close lArqNum
iErrosPnt = 56
    
    '############################################
    'Inserido por Wagner 16/05/2006
    'Verifica se existem tsks faltando ou desatualizado
    lErro = Verifica_Tsks(iCarregaForm)
iErrosPnt = 57
    If lErro <> SUCESSO Then Error 6000
iErrosPnt = 58

    lErro = Verifica_SubPastas(iCarregaForm)
    If lErro <> SUCESSO Then Error 6000
    '############################################
    
    If iCarregaForm Then
iErrosPnt = 59
     
       Atualiza.Show vbModal
iErrosPnt = 60
       
    Else
iErrosPnt = 61
    
       For iIndice = 1 To colOutrosArquivos.Count
iErrosPnt = 62
       
            sMsgErroAux = "Tipo: Out Linha: " & CStr(iIndice)
iErrosPnt = 63
       
           sNomeArq = colOutrosArquivos.Item(iIndice)
iErrosPnt = 64
           
            sMsgErroAux = sMsgErroAux & " NomeArq: " & sNomeArq
iErrosPnt = 65
           
           If Len(Dir(sDirOrigem & sNomeArq)) = 0 Then Error 5008
iErrosPnt = 66
       
           dDataArqOrigem = FileDateTime(sDirOrigem & sNomeArq)
iErrosPnt = 67
        
           If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
iErrosPnt = 68
           
              iCarregaForm = True
iErrosPnt = 69
              Exit For
iErrosPnt = 70
        
           Else
iErrosPnt = 71
        
              If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
iErrosPnt = 72
                 iCarregaForm = True
iErrosPnt = 73
                 Exit For
iErrosPnt = 74
              End If
iErrosPnt = 75
              
              dDataArqDestino = FileDateTime(sDirDestino & sNomeArq)
iErrosPnt = 76
          
              If dDataArqOrigem > dDataArqDestino Then
iErrosPnt = 77
                 iCarregaForm = True
iErrosPnt = 78
                 Exit For
iErrosPnt = 79
              End If
iErrosPnt = 80
        
           End If
iErrosPnt = 81
     
       Next iIndice
iErrosPnt = 82
       
       If iCarregaForm Then
iErrosPnt = 83
         
          Atualiza.Show vbModal
iErrosPnt = 84
       
       End If
iErrosPnt = 85
            
    End If
iErrosPnt = 86
      
    Verifica_Atualizacoes = SUCESSO
    
    Exit Function
    
Erro_Verifica_Atualizacoes:

    Verifica_Atualizacoes = Err.Number

    Beep
    
    Select Case Err.Number
    
       Case Is = 5005
          MsgBox "Não há informação sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
       Case Is = 5006
          MsgBox "Informação incorreta sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
       Case Is = 5007
          MsgBox "Não há informação sobre os arquivos de atualização no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
       Case Is = 5008
          MsgBox "O arquivo " & Chr(34) & sDirOrigem & sNomeArq & Chr(34) & " não existe.", vbCritical, "Atenção!"
       
       Case 6000, 9999
       
        Case 10002
            MsgBox "Erro: " & Err.Number & " - O diretório " & sDirDestino & " está com restrições ao acesso.", vbCritical, "Atenção!"
       
       Case Else
         MsgBox Err.Number & "-" & Err.Description & ". " & CStr(iErrosPnt) & "-" & sMsgErro & " " & sMsgErroAux, vbCritical, "Atenção!"

    End Select
    
    Exit Function

End Function

'##################################
'Inserido por Wagner 16/05/2006
Public Function Verifica_Tsks(iCarregaForm As Integer) As Long

Dim lErro As Long
Dim sDirTskSRV As String
Dim sDirTskPRG As String
Dim objFileSRV As FileListBox
Dim objFilePRG As FileListBox
Dim iIndSRV As Integer
Dim iIndPRG As Integer
Dim bTemTsk As Boolean
Dim bTskDataAnt As Boolean
Dim sMsgErro As String, sMsgErroAux As String, iLinha As Integer, iErrosPnt As Integer

On Error GoTo Erro_Verifica_Tsks

    sMsgErro = "Função: Verifica_Tsks"
iErrosPnt = 1

    'Obtém o diretório dos tsk do servidor
    sDirTskSRV = LeArqINI("Geral", "Dirtsk", NOME_ARQUIVO_ADM)
iErrosPnt = 2
    If Len(sDirTskSRV) = 0 Then Error 6001
iErrosPnt = 3
    
    sMsgErro = sMsgErro & " Origem: " & sDirTskSRV
iErrosPnt = 4
    
    'Obtém o diretório dos tsk das máquinas
    sDirTskPRG = LeArqINI("ForPrint", "Dirtsks", NOME_ARQUIVO_ADM)
iErrosPnt = 5
    If Len(sDirTskPRG) = 0 Then Error 6002
iErrosPnt = 6
    
    sMsgErro = sMsgErro & " Destino: " & sDirTskPRG
iErrosPnt = 7

    lErro = Teste_Acesso_Full_Pasta(sDirTskPRG)
    If lErro <> SUCESSO Then Error 10002
    
    If UCase(sDirTskSRV) <> UCase(sDirTskPRG) Then
iErrosPnt = 8
    
        Set objFileSRV = Atualiza.File1
iErrosPnt = 9
        Set objFilePRG = Atualiza.File2
iErrosPnt = 10
        
        objFileSRV.Path = sDirTskSRV
iErrosPnt = 11
        objFilePRG.Path = sDirTskPRG
iErrosPnt = 12
        
        'Para cada arquivo no servidor
        For iIndSRV = 0 To objFileSRV.ListCount - 1
iErrosPnt = 13
        
            bTemTsk = False
iErrosPnt = 14
            bTskDataAnt = False
iErrosPnt = 15
            sMsgErroAux = "NomeArqSrv: " & objFileSRV.List(iIndSRV)
iErrosPnt = 16
            
            'Para cada arquivo na máquina
            For iIndPRG = 0 To objFilePRG.ListCount - 1
iErrosPnt = 17
            
                sMsgErroAux = "NomeArqSrv: " & objFileSRV.List(iIndSRV) & "NomeArqPgm: " & objFilePRG.List(iIndPRG)
iErrosPnt = 18
            
                'Se o arquivo for o mesmo
                If UCase(objFileSRV.List(iIndSRV)) = UCase(objFilePRG.List(iIndPRG)) Then
iErrosPnt = 19
                    bTemTsk = True
iErrosPnt = 20
                    'Se as data do tsk da máquina é anterior a do servidor
                    If FileDateTime(objFileSRV.Path & "\" & objFileSRV.List(iIndSRV)) > FileDateTime(objFilePRG.Path & "\" & objFilePRG.List(iIndPRG)) Then
iErrosPnt = 21
                        bTskDataAnt = True
iErrosPnt = 22
                    End If
iErrosPnt = 23
                    Exit For
iErrosPnt = 24
                End If
iErrosPnt = 25
            
            Next
iErrosPnt = 26
            
            'Se não existe o arquivo na máquina ou o arquivo do servidor é mais recente
            If bTskDataAnt Or Not bTemTsk Then
iErrosPnt = 27
                iCarregaForm = True
iErrosPnt = 28
                Exit For
iErrosPnt = 29
            End If
iErrosPnt = 30
            
        Next
iErrosPnt = 31
        
    End If
iErrosPnt = 32
        
    Verifica_Tsks = SUCESSO
   
    Exit Function

Erro_Verifica_Tsks:

    Verifica_Tsks = Err
   
    Select Case Err
    
        Case 6001, 6002
            MsgBox "Não há informação sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
   
        Case Else
            MsgBox Err.Number & "-" & Err.Description & ". " & CStr(iErrosPnt) & "-" & sMsgErro & " " & sMsgErroAux, vbCritical, "Atenção!"

    End Select
   
   Exit Function
   
End Function


Public Function Verifica_SubPastas(iCarregaForm As Integer) As Long

Dim lErro As Long
Dim sDirOrigem As String
Dim sDirDestino As String, sDirDestinoAux As String
Dim objSubPastaOrigem As DirListBox
Dim objSubPastaDestino As DirListBox
Dim iIndOrigem As Integer
Dim iIndDestino As Integer, bTemSubDir As Boolean
Dim sMsgErro As String, sMsgErroAux As String, iLinha As Integer
Dim sComputador As String

On Error GoTo Erro_Verifica_SubPastas

    sMsgErro = "Função: Verifica_SubPastas"

    'Obtém o diretório dos tsk do servidor
    sDirOrigem = LeArqINI("Geral", "DirProgram", NOME_ARQUIVO_ADM)
    If Len(sDirOrigem) = 0 Then Error 6001
    
    sMsgErro = sMsgErro & " Origem: " & sDirOrigem
    
    sDirDestino = App.Path
    If Right(sDirDestino, 1) <> "\" Then
       sDirDestino = sDirDestino & "\"
    End If
    
    sMsgErro = sMsgErro & " Destino: " & sDirDestino

    lErro = Teste_Acesso_Full_Pasta(sDirDestino)
    If lErro <> SUCESSO Then Error 10002
    
    sComputador = String(512, 0)
    Call GetComputerName(sComputador, Len(sComputador))
    sComputador = Replace(sComputador, Chr(0), "")
    
    sDirDestinoAux = ""
    sDirDestinoAux = "\\" & sComputador & Mid(sDirDestino, 3)
   
    If UCase(sDirOrigem) <> UCase(sDirDestino) And UCase(sDirOrigem) <> UCase(sDirDestinoAux) Then
    
        Set objSubPastaOrigem = Atualiza.Dir1
        Set objSubPastaDestino = Atualiza.Dir2
        
        objSubPastaOrigem.Path = sDirOrigem
        objSubPastaDestino.Path = sDirDestino
        
        'Para cada subpasta no servidor
        For iIndOrigem = 0 To objSubPastaOrigem.ListCount - 1
        
            bTemSubDir = False
             sMsgErroAux = "NomeSubDirOrig: " & objSubPastaOrigem.List(iIndOrigem)
            
            'Para cada subpasta
            For iIndDestino = 0 To objSubPastaDestino.ListCount - 1
            
                sMsgErroAux = "NomeSubDirOrig: " & objSubPastaOrigem.List(iIndOrigem) & " NomeSubDirDest: " & objSubPastaDestino.List(iIndDestino)
           
                'Se a subpasta for a mesmo
                If Replace(UCase(objSubPastaOrigem.List(iIndOrigem)), UCase(sDirOrigem), "") = Replace(UCase(objSubPastaDestino.List(iIndDestino)), UCase(sDirDestino), "") Then
                    bTemSubDir = True
                    Exit For
                End If
            
            Next
            
            'Se não existe o subdiretório
            If Not bTemSubDir Then
                iCarregaForm = True
                Exit For
            End If
            
        Next
       
    End If
        
    Verifica_SubPastas = SUCESSO
   
    Exit Function

Erro_Verifica_SubPastas:

    Verifica_SubPastas = Err
   
    Select Case Err
    
        Case 6001, 6002
            MsgBox "Não há informação sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
   
        Case Else
            MsgBox Err.Number & "-" & Err.Description & ". " & sMsgErro & " " & sMsgErroAux, vbCritical, "Atenção!"

    End Select
   
   Exit Function
   
End Function
'#################################

Public Function Testa_Arquivo(ByVal sNomeArq As String) As Long

Dim dtData As Date

On Error GoTo Erro_Testa_Arquivo

    dtData = FileDateTime(sNomeArq)

    Testa_Arquivo = SUCESSO
   
    Exit Function

Erro_Testa_Arquivo:

    Testa_Arquivo = Err
   
    Select Case Err
    
        Case Else
            MsgBox "Erro: " & Err.Number & " - Ao tentar acessar o arquivo " & sNomeArq & ". Detalhe: " & Err.Description & ").", vbCritical, "Atenção!"

    End Select
   
   Exit Function
   
End Function

Public Function Copia_Arquivo(ByVal sDirO As String, ByVal sDirD As String, ByVal sNomeArq As String) As Long

Dim lErro As Long
Dim sNomeArqO As String
Dim sNomeArqD As String
Dim iPos As Integer
Dim sDiretorio As String

On Error GoTo Erro_Copia_Arquivo

    sNomeArqO = sDirO & sNomeArq
    sNomeArqD = sDirD & sNomeArq
    
    sDiretorio = sDirO
    If Not GetAttr(sDiretorio) And vbDirectory Then Error 10000
    
    sDiretorio = sDirD
    If Not GetAttr(sDiretorio) And vbDirectory Then Error 10001
    
    lErro = Copia_Arquivo1(sDirO, sDirD, sNomeArq)
    If lErro <> SUCESSO Then
        Sleep (10000)
        lErro = Copia_Arquivo1(sDirO, sDirD, sNomeArq)
        If lErro <> SUCESSO Then Error 10003
    End If

    Copia_Arquivo = SUCESSO
   
    Exit Function

Erro_Copia_Arquivo:

    Copia_Arquivo = Err
   
    Select Case Err
    
        Case 52, 53, 10000, 10001
            MsgBox "Erro: " & Err.Number & " - O diretório " & sDiretorio & " ou o acesso a ele foi negado. Detalhe: " & Err.Description & ").", vbCritical, "Atenção!"
    
'        Case 10002
'            MsgBox "Erro: " & Err.Number & " - O diretório " & sDiretorio & " está com restrições ao acesso.", vbCritical, "Atenção!"
    
        Case 10003
            MsgBox "Erro: " & Err.Number & " - Ao copiar o arquivo " & sNomeArqO & " para " & sNomeArqD & ". Certifique-se que ele não está em uso.", vbCritical, "Atenção!"
    
        Case Else
            MsgBox "Erro: " & Err.Number & " - Ao copiar o arquivo " & sNomeArqO & " para " & sNomeArqD & ". Detalhe: " & Err.Description & ").", vbCritical, "Atenção!"

    End Select
   
   Exit Function
   
End Function

Private Function Copia_Arquivo1(ByVal sDirO As String, ByVal sDirD As String, ByVal sNomeArq As String) As Long
On Error GoTo Erro_Copia_Arquivo1
    
    Call WinExec("attrib -r " & sDirD & " /D /S", SW_HIDE)
    
    FileCopy sDirO & sNomeArq, sDirD & sNomeArq
    Copia_Arquivo1 = SUCESSO
    Exit Function
Erro_Copia_Arquivo1:
    Copia_Arquivo1 = Err
End Function

Public Function Teste_Acesso_Full_Pasta(ByVal sDir As String) As Long
Dim sArq As String
On Error GoTo Erro_Teste_Acesso_Full_Pasta
    sArq = Dir(sDir & "TesteAcesso.txt")
    If sArq = "TesteAcesso.txt" Then
       Kill sDir & "TesteAcesso.txt"
    End If
    Open sDir & "TesteAcesso.txt" For Output As #100
    Print #100, "Teste"
    Close #100
    Kill sDir & "TesteAcesso.txt"
    Teste_Acesso_Full_Pasta = SUCESSO
    Exit Function
Erro_Teste_Acesso_Full_Pasta:
    Teste_Acesso_Full_Pasta = Err
    Close #100
End Function

Public Function Copia_Pasta(ByVal sDirO As String, ByVal sDirD As String) As Long
On Error GoTo Erro_Copia_Pasta
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Call SetFileAttributes(sDirD, FILE_ATTRIBUTE_NORMAL)
    Call WinExec("attrib -r " & sDirD & "/*.* /D /S", SW_HIDE)
    
    FSO.CopyFolder sDirO, sDirD
    Copia_Pasta = SUCESSO
    Exit Function
Erro_Copia_Pasta:
    Copia_Pasta = Err
End Function
