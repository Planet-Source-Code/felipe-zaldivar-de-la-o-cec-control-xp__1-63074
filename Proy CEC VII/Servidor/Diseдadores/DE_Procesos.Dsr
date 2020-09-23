VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DE_Procesos 
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10530
   _ExtentX        =   18574
   _ExtentY        =   15849
   FolderFlags     =   1
   TypeLibGuid     =   "{4244B051-1361-43D0-9140-06D7E400AB05}"
   TypeInfoGuid    =   "{676466E9-1B2B-4223-936F-62ED6752A55D}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "ConPrincipal"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Mode=Read;Persist Security Info=False;Jet OLEDB:Database Password=control"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "Tbl_Acceso"
      CommDispId      =   1010
      RsDispId        =   1015
      CommandText     =   "Tbl_Acceso"
      ActiveConnectionName=   "ConPrincipal"
      CommandType     =   2
      dbObjectType    =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "C_Grupo"
         Caption         =   "C_Grupo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Nombre"
         Caption         =   "Nombre"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   202
         Name            =   "C_Acceso"
         Caption         =   "C_Acceso"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "Password"
         Caption         =   "Password"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Nivel"
         Caption         =   "Nivel"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Amonestaciones"
         Caption         =   "Amonestaciones"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Usr_Bloqueado"
         Caption         =   "Usr_Bloqueado"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "N_Bloqueos"
         Caption         =   "N_Bloqueos"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Fecha_Reg"
         Caption         =   "Fecha_Reg"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   202
         Name            =   "C_U_Registro"
         Caption         =   "C_U_Registro"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Tbl_Procesos_Reg"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Tbl_Procesos_Reg where proceso Like '%Bloqueado%' order by fecha desc, fecha desc"
      ActiveConnectionName=   "ConPrincipal"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Tbl_Acceso"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   202
         Name            =   "C_Acceso"
         Caption         =   "C_Acceso"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "C_Maq"
         Caption         =   "C_Maq"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Fecha"
         Caption         =   "Fecha"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Hora"
         Caption         =   "Hora"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   250
         Scale           =   0
         Type            =   202
         Name            =   "Proceso"
         Caption         =   "Proceso"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "C_Acceso"
         ChildField      =   "C_Acceso"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DE_Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
