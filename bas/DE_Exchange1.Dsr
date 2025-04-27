VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   13860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   _ExtentX        =   33761
   _ExtentY        =   24448
   FolderFlags     =   1
   TypeLibGuid     =   "{EAA92514-A95A-4FC2-AC6E-0E66CEB13D0D}"
   TypeInfoGuid    =   "{4F45D7AD-9193-4F04-80F5-A3CB1E9019C1}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "ExchConn"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=PBKS;Data Source=PAPYRUS-94TNP9S"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "Exchanges"
      CommDispId      =   1002
      RsDispId        =   1029
      CommandText     =   "dbo.tExchange"
      ActiveConnectionName=   "ExchConn"
      CommandType     =   2
      dbObjectType    =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   72
         Name            =   "EXCH_ID"
         Caption         =   "EXCH_ID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   72
         Name            =   "EXCH_ZSessionID"
         Caption         =   "EXCH_ZSessionID"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   72
         Name            =   "EXCH_OPSESSIONID"
         Caption         =   "EXCH_OPSESSIONID"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "EXCH_Date"
         Caption         =   "Time"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EXCH_Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "EXCH_TillCode"
         Caption         =   "EXCH_TillCode"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EXCH_OperatorID"
         Caption         =   "EXCH_OperatorID"
      EndProperty
      BeginProperty Field8 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "EXCH_General_Disc"
         Caption         =   "EXCH_General_Disc"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EXCH_CostValue"
         Caption         =   "EXCH_CostValue"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EXCH_SupervisorID"
         Caption         =   "EXCH_SupervisorID"
      EndProperty
      BeginProperty Field11 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EXCH_ChangeGiven"
         Caption         =   "EXCH_ChangeGiven"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Products"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"DE_Exchange1.dsx":0000
      ActiveConnectionName=   "ExchConn"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Exchanges"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "P_Title"
         Caption         =   "Description"
         Control         =   "Microsoft Hierarchical FlexGrid Control 6.0 (SP4) (OLEDB)"
         ControlGuid     =   "{0ECD9B64-23AA-11D0-B351-00A0C9055D8E}"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Price"
         Caption         =   "Price"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CSL_Qty"
         Caption         =   "Qty"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "CSL_TimeOfSale"
         Caption         =   "Time"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   72
         Name            =   "CSL_Exchange_GUID"
         Caption         =   "CSL_Exchange_GUID"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Payments"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT * FROM tPAYMENT"
      ActiveConnectionName=   "ExchConn"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Exchanges"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "PAY_ID"
         Caption         =   "PAY_ID"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "PAY_Amt"
         Caption         =   "Amt"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   72
         Name            =   "PAY_Exchange_GUID"
         Caption         =   "PAY_Exchange_GUID"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "PAY_Amt_Tendered"
         Caption         =   "PAY_Amt_Tendered"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "PAY_PaymentType"
         Caption         =   "Type"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "PAY_CCExpiryDate"
         Caption         =   "PAY_CCExpiryDate"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "PAY_CCLastFour"
         Caption         =   "PAY_CCLastFour"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "PAY_Tot_Received"
         Caption         =   "PAY_Tot_Received"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
