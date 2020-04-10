VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalUniao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ATENÇÃO!!! NÃO EXCLUIR O CÓDIGO COMENTADO ABAIXO

'Public objContabil As ClassContabil             'Todas
'Public gobjTribTab As ClassTribTab              'Todas
'
'Public objGrid1 As AdmGrid                      'Todas
'Public objGridItens As AdmGrid                  'Todas
'
'Public giFilialEntregaManual As Integer         'Todas
'Public giFilialEntregaAlterada As Integer       'Todas
'Public iValorIRRFAlterado As Integer            'Todas
'Public iValorDespesasAlterado As Integer        'Todas
'Public iValorSeguroAlterado As Integer          'Todas
'Public iValorFreteAlterado As Integer           'Todas
'Public iValorDescontoAlterado As Integer        'Todas
'Public iAlterado As Integer                     'Todas
'Public giPosCargaOk As Integer                  'Todas

Public iClienteAlterado As Integer              'NF , NFF,    ,    , NFD, NFR
Public objGridAlocacoes As AdmGrid              'NF, NFF, NFFP, NFP
Public objGridComissoes As AdmGrid              'NF, NFF, NFFP, NFP
Public objTabComissoes As ClassTabComissoes     'NF, NFF, NFFP, NFP
Public gdDesconto As Double                     'NF, NFF, NFFP, NFP

Public objGridParcelas As AdmGrid               '  , NFF, NFFP
Public giDataReferenciaAlterada As Integer      '  , NFF, NFFP
Public iPercAcresFinAlterado As Integer         '  , NFF, NFFP

Public iFornecedorAtivo As Integer              '  ,    ,     ,    , NFD, NFR
Public iFornecedorAlterado As Integer           '  ,    ,     ,    , NFD, NFR

Public glCodPedido As Long                      '  ,    , NFFP, NFP
Public giFilialPedido As Integer                '  ,    , NFFP, NFP
Public giFaturaIntegral As Integer              '  ,    , NFFP, NFP
Public gcolItemPedido As New Collection         '  ,    , NFFP, NFP
Public gcolItemQuant As New Collection          '  ,    , NFFP, NFP
