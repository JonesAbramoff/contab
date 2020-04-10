VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRBGravaArtmill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Tributacao_Iniciar(ByVal colDriversTributos As Collection) As Long

Dim lErro As Long
Dim objTributoDriver As ClassTributoDriver

On Error GoTo Erro_Tributacao_Iniciar

    Set objTributoDriver = New ClassTributoISSArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27247
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoIRArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27248
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoIPIArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27249
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoICMSArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27250
    colDriversTributos.Add objTributoDriver
        
    Set objTributoDriver = New ClassTributoINSSArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
        
    'valendo a partir de 01/12/2002 pela mp 66
    Set objTributoDriver = New ClassTributoPISArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    'valendo a partir de 01/02/2004 pela mp 135
    Set objTributoDriver = New ClassTributoCOFINSArtmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    'valendo a partir de 01/02/2004 pela mp 135
    Set objTributoDriver = New ClassTribArt30Artmill
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    Tributacao_Iniciar = SUCESSO
     
    Exit Function
    
Erro_Tributacao_Iniciar:

    Tributacao_Iniciar = gErr
     
    Select Case gErr
          
        Case 27247 To 27250, 32251
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153654)
     
    End Select
     
    Exit Function

End Function

