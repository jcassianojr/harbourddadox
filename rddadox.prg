
#include "rddsys.ch"
#include "fileio.ch"
#include "error.ch"
#include "rddadox.ch"
#include "dbstruct.ch"
#include "dbinfo.ch"

#include "hbusrrdd.ch"

#define WA_RECORDSET   1
#define WA_BOF         2
#define WA_EOF         3
#define WA_CONNECTION  4
#define WA_CATALOG     5
#define WA_TABLENAME   6
#define WA_ENGINE      7
#define WA_SERVER      8
#define WA_USERNAME    9
#define WA_PASSWORD   10
#define WA_QUERY      11
#define WA_LOCATEFOR  12
#define WA_SCOPEINFO  13
#define WA_SQLSTRUCT  14
#define WA_CONNOPEN   15
#define WA_PENDINGREL 16
#define WA_FOUND      17
#define WA_FILTERBLOCK 18

#define WA_SIZE       18


#define RDD_CONNECTION 1
#define RDD_CATALOG    2

#define RDD_SIZE       2

#define UR_TRANSBEGIN     51
#define UR_TRANSCOMMIT    52
#define UR_TRANSROLLBACK  53
#define UR_GETROW         54
#define UR_GETROWBLANK    55
#define UR_PUTROW         56

ANNOUNCE RDDADOX

THREAD STATIC t_cTableName
THREAD STATIC t_cEngine
THREAD STATIC t_cServer
THREAD STATIC t_cUserName
THREAD STATIC t_cPassword
THREAD STATIC t_cQuery := ""


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_INIT()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_INIT( nRDD )

   LOCAL aRData := Array( RDD_SIZE )

   USRRDD_RDDDATA( nRDD, aRData )

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_NEW()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_NEW( nWA )

   LOCAL aWAData := Array( WA_SIZE )

   aWAData[ WA_BOF ] := .F.
   aWAData[ WA_EOF ] := .F.

   USRRDD_AREADATA( nWA, aWAData )

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +    Static Function ADO_CREATEFIELDS()
// +    Apenas guarda a matriz de estrutura para ser processada no CREATE
// +
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_CREATEFIELDS( nWA, aStruct )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   
   // Em vez de montar uma string ruim aqui, guardamos o Array original 
   // para o dialeto processar na ADO_CREATE
   aWAData[ WA_SQLSTRUCT ] := aStruct

   RETURN HB_SUCCESS

// +--------------------------------------------------------------------
// +
// +    Static Function ADO_CREATE()
// +
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_CREATE( nWA, aOpenInfo )
   LOCAL cDataBase  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 1, ";" )
   LOCAL cTableName := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 2, ";" )
   LOCAL cDbEngine  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 3, ";" )
   LOCAL cServer    := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 4, ";" )
   LOCAL cUserName  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 5, ";" )
   LOCAL cPassword  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 6, ";" )

   LOCAL oConnection := win_oleCreateObject( "ADODB.Connection" )
   LOCAL oCatalog    := win_oleCreateObject( "ADOX.Catalog" )
   LOCAL aWAData     := USRRDD_AREADATA( nWA )
   LOCAL oError, n, cSqlTable

   LOCAL cEXTENSAO
   LOCAL cDIRETORIO
   LOCAL cNAME

   cNAME      := ""
   cEXTENSAO  := ""
   cDIRETORIO := ""

   hb_FNameSplit( cDataBase, @cDIRETORIO, @cName, @CEXTENSAO )
   cEXTENSAO := Lower( cEXTENSAO )
   cDbEngine := Upper( cDbEngine )

   // (O bloco DO CASE de conexão original permanece aqui até fazermos o item 4)
   DO CASE
   CASE cEXTENSAO == ".accdb" .AND. cDbEngine == "ACCDB"
      IF !hb_FileExists( cDataBase )
         oCatalog:Create( "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + cDataBase )
      ENDIF
      oConnection:Open( "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + cDataBase )
   CASE cDbEngine == "SQLITE" .OR. cEXTENSAO == ".sqlite" .OR. cEXTENSAO == ".sqlite3" .OR. cEXTENSAO == ".fossil" .OR. cEXTENSAO == ".db3"
      IF !hb_FileExists( cDataBase )
         oCatalog:Create( "DRIVER=SQLite3 ODBC Driver;Database=" + cDataBase )
      ENDIF
      oConnection:Open( "DRIVER=SQLite3 ODBC Driver;Database=" + cDataBase )
   CASE cEXTENSAO == ".mdb" .OR. cDbEngine == "ACCESS"
      IF !hb_FileExists( cDataBase )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )
   CASE cEXTENSAO == ".xls" .OR. cDbEngine == "XLS"
      IF !hb_FileExists( cDataBase )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )
   CASE cEXTENSAO == ".db" .OR. cDbEngine == "PARADOX"
      IF Right( AllTrim( cDIRETORIO ), 1 ) != "\" .AND. Right( AllTrim( cDIRETORIO ), 1 ) != "/"
        cDIRETORIO += "\"
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDIRETORIO + ";Extended Properties='Paradox 5.x';" )
   CASE cEXTENSAO == ".fdb" .OR. cEXTENSAO == ".gdb" .OR. cEXTENSAO == ".ib" .OR. cDbEngine == "FIREBIRD"
      IF !hb_FileExists( cDataBase )
         oCatalog:Create( "Driver=Firebird ODBC Driver;Uid=" + cUserName + ";Pwd=" + cPassword + ";DbName=" + cDataBase + ";" )
      ENDIF
      oConnection:Open( "Driver=Firebird ODBC Driver;Uid=" + cUserName + ";Pwd=" + cPassword + ";DbName=" + cDataBase + ";" )
      oConnection:CursorLocation := adUseClient
   CASE cDbEngine == "MARIADB" 
      oConnection:Open( "Driver={MariaDB ODBC 3.2 Driver};server=" + cServer + ";database=" + cDataBase + ";uid=" + cUserName + ";pwd=" + cPassword )
   CASE cDbEngine == "PGSQL" .OR. cDbEngine == "POSTGRESQL" 
      oConnection:Open( "Driver={PostgreSQL ANSI};server=" + cServer + ";database=" + cDataBase + ";uid=" + cUserName + ";pwd=" + cPassword )
   CASE cDbEngine == "PGSQL64" 
      oConnection:Open( "Driver={PostgreSQL ANSI(x64)};server=" + cServer + ";database=" + cDataBase + ";uid=" + cUserName + ";pwd=" + cPassword )
   CASE cDbEngine == "MYSQL" 
      oConnection:Open( "Driver={MySQL ODBC 8.0 ANSI Driver};server=" + cServer + ";database=" + cDataBase + ";uid=" + cUserName + ";pwd=" + cPassword )
   CASE cDbEngine == "MYSQL64" 
      oConnection:Open( "Driver={MySQL ODBC 9.0 ANSI Driver};server=" + cServer + ";database=" + cDataBase + ";uid=" + cUserName + ";pwd=" + cPassword )
   ENDCASE

   // Monta a estrutura da tabela usando o Dialeto Avançado
   cSqlTable := ADO_BUILD_CREATE_TABLE( cTableName, aWAData[ WA_SQLSTRUCT ], cDbEngine )

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oConnection:Execute( "DROP TABLE " + cTableName )
   RECOVER
   END SEQUENCE

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      // Executa o comando de criação perfeitamente adaptado
      oConnection:Execute( cSqlTable )
   RECOVER
      oError             := ErrorNew()
      oError:GenCode     := EG_CREATE
      oError:SubCode     := 1004
      oError:Description := hb_langErrMsg( EG_CREATE ) + " (" + hb_langErrMsg( EG_UNSUPPORTED ) + ")"
      oError:FileName   := aOpenInfo[ UR_OI_NAME ]
      oError:CanDefault := .T.
      FOR n := 0 TO oConnection:Errors:COUNT - 1
         oError:Description += hb_eol() + oConnection:Errors( n ) :Description
      NEXT
      UR_SUPER_ERROR( nWA, oError )
   END SEQUENCE

   oConnection:Close()
   RETURN HB_SUCCESS

// +--------------------------------------------------------------------
// +    Static Function ADO_BUILD_CREATE_TABLE()
// +    Construção inteligente da tabela com tipos nativos e SR_RECNO
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_BUILD_CREATE_TABLE( cTableName, aStruct, cEngine )
   LOCAL cSql := ""
   LOCAL i, mFldNm, mFldType, mFldLen, mFldDec

   // Proteção da sintaxe IF NOT EXISTS
   IF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB" .OR. ;
      cEngine == "FIREBIRD" .OR. cEngine == "ORACLE" .OR. cEngine == "OCI" .OR. ;
      cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cSql := "CREATE TABLE " + cTableName + " ( "
   ELSE
      cSql := "CREATE TABLE IF NOT EXISTS " + cTableName + " ( "
   ENDIF

   // --- INSERÇÃO OBRIGATÓRIA DOS CAMPOS DE CONTROLE DO HARBOUR ---
   // 1. SR_RECNO (Auto-incremento nativo por banco)
   DO CASE
   CASE cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
      cSql += " SR_RECNO INT NOT NULL AUTO_INCREMENT UNIQUE, "
   CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
      cSql += " SR_RECNO NUMBER(10,0) GENERATED ALWAYS AS IDENTITY UNIQUE, "
   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cSql += " SR_RECNO SERIAL UNIQUE, "
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cSql += " SR_RECNO INT IDENTITY(1,1) UNIQUE, "
   CASE cEngine == "SQLITE"
      cSql += " SR_RECNO INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE, "
   CASE cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
      cSql += " SR_RECNO COUNTER UNIQUE, "
   CASE cEngine == "FIREBIRD" .OR. cEngine == "FDB" .OR. cEngine == "GDB" .OR. cEngine == "IB"
      cSql += " SR_RECNO INTEGER GENERATED ALWAYS AS IDENTITY UNIQUE, "
   CASE cEngine == "DUCKDB" .OR. cEngine == "DUCKLAKE"
      cSql += " SR_RECNO BIGINT, "
   OTHERWISE
      cSql += " SR_RECNO INT UNIQUE, "
   ENDCASE

   // 2. SR_DELETED (Campo lógico de exclusão)
   cSql += " SR_DELETED CHAR(1) DEFAULT ' ' NOT NULL, "

   // --- MAPEAMENTO DOS CAMPOS DO USUÁRIO ---
   FOR i := 1 TO Len( aStruct )
      mFldNm   := aStruct[ i, DBS_NAME ]
      mFldType := aStruct[ i, DBS_TYPE ]
      mFldLen  := aStruct[ i, DBS_LEN ]
      mFldDec  := aStruct[ i, DBS_DEC ]

      // Pula se o usuário já tiver definido os campos de controle no DBSTRUCT
      IF Upper( mFldNm ) == "SR_RECNO" .OR. Upper( mFldNm ) == "SR_DELETED"
         LOOP
      ENDIF

      // Ajuste Case Sensitive para PostgreSQL
      IF cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
         mFldNm := Upper( mFldNm )
      ENDIF

      cSql += mFldNm + " "

      DO CASE
      // CARACTER
      CASE mFldType == "C"
         IF cEngine == "ORACLE" .OR. cEngine == "OCI"
            cSql += "VARCHAR2(" + hb_ntos( mFldLen ) + ")"
         ELSEIF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL" .OR. ;
                cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
            cSql += "VARCHAR(" + hb_ntos( mFldLen ) + ")"
         ELSEIF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
            cSql += "VARCHAR(" + hb_ntos( mFldLen ) + ") DEFAULT ''"
         ELSEIF cEngine == "SQLITE"
            cSql += "TEXT NOT NULL DEFAULT ''"
         ELSEIF cEngine == "FIREBIRD"
            cSql += "VARCHAR(" + hb_ntos( mFldLen ) + ")"
         ELSE
            cSql += "CHAR(" + hb_ntos( mFldLen ) + ")"
         ENDIF

      // NUMERICO
      CASE mFldType == "N"
         IF cEngine == "ORACLE" .OR. cEngine == "OCI"
            cSql += "NUMBER(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ") DEFAULT 0"
         ELSEIF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL" .OR. ;
                cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
            IF mFldDec > 0
               cSql += "NUMERIC(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
            ELSE
               cSql += iif( mFldLen <= 9, "INT", "BIGINT" )
            ENDIF
         ELSEIF cEngine == "FIREBIRD"
            IF mFldDec > 0
               cSql += "DECIMAL(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
            ELSE
               cSql += iif( mFldLen <= 4, "SMALLINT", iif( mFldLen <= 9, "INTEGER", "BIGINT" ) )
            ENDIF
         ELSEIF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
            cSql += iif( mFldDec > 0, "DOUBLE DEFAULT 0", "LONG DEFAULT 0" )
         ELSEIF cEngine == "SQLITE"
            cSql += "NUMERIC(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ") DEFAULT 0"
         ELSEIF cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
            IF mFldDec > 0
               cSql += "NUMERIC(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
            ELSE
               cSql += iif( mFldLen <= 9, "INTEGER", "BIGINT" )
            ENDIF
         ELSE
            cSql += "NUMERIC(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
         ENDIF

      // DATA E HORA
      CASE mFldType == "D"
         IF cEngine == "SQLITE"
            cSql += "DATE NOT NULL DEFAULT ''"
         ELSEIF cEngine == "ORACLE" .OR. cEngine == "OCI" .OR. cEngine == "FIREBIRD"
            cSql += "TIMESTAMP"
         ELSEIF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL" .OR. ;
                cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB" .OR. ;
                cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
            cSql += "DATETIME"
         ELSE
            cSql += "DATE"
         ENDIF

      CASE mFldType == "@" .OR. mFldType == "T"
         IF cEngine == "ORACLE" .OR. cEngine == "OCI" .OR. ;
            cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL" .OR. ;
            cEngine == "FIREBIRD"
            cSql += "TIMESTAMP"
         ELSE
            cSql += "DATETIME"
         ENDIF

      // LOGICO
      CASE mFldType == "L"
         IF cEngine == "ORACLE" .OR. cEngine == "OCI"
            cSql += "SMALLINT"
         ELSEIF cEngine == "FIREBIRD"
            cSql += "SMALLINT DEFAULT 0 NOT NULL"
         ELSEIF cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL" .OR. cEngine == "SQLITE"
            cSql += "BOOLEAN"
         ELSEIF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB" .OR. ;
                cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
            cSql += "BIT DEFAULT 0"
         ELSE
            cSql += "BOOL"
         ENDIF

      // MEMO E BINARIOS
      CASE mFldType == "M"
         IF cEngine == "ORACLE" .OR. cEngine == "OCI"
            cSql += "CLOB"
         ELSEIF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL" .OR. ;
                cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
            cSql += "TEXT"
         ELSEIF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
            cSql += "LONGTEXT"
         ELSEIF cEngine == "FIREBIRD"
            cSql += "BLOB SUB_TYPE TEXT"
         ELSE
            cSql += "TEXT"
         ENDIF

      CASE mFldType == "G" .OR. mFldType == "P" .OR. mFldType == "W"
         IF cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
            cSql += "BYTEA"
         ELSEIF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
            cSql += "VARBINARY(MAX)"
         ELSEIF cEngine == "MDB" .OR. cEngine == "ACCESS" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
            cSql += "OLEOBJECT"
         ELSEIF cEngine == "FIREBIRD"
            cSql += "BLOB SUB_TYPE 0"
         ELSE
            cSql += "BLOB"
         ENDIF

      // DOUBLE/FLOAT
      CASE mFldType == "B" .OR. mFldType == "F"
         IF cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
            cSql += "FLOAT"
         ELSEIF cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
            cSql += "DOUBLE(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
         ELSEIF cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
            cSql += "NUMERIC(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
         ELSEIF cEngine == "FIREBIRD"
            cSql += "DECIMAL(" + hb_ntos( mFldLen ) + "," + hb_ntos( mFldDec ) + ")"
         ELSE
            cSql += "DOUBLE"
         ENDIF

      OTHERWISE
         cSql += "VARCHAR(" + hb_ntos( mFldLen ) + ")"
      ENDCASE

      IF i < Len( aStruct )
         cSql += ", "
      ENDIF
   NEXT

   cSql += " )"
   
   RETURN cSql



// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_CLOSE()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_CLOSE( nWA )

   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:Close()
      IF !Empty( aWAData[ WA_CONNOPEN ] )
         IF aWAData[ WA_CONNECTION ] :STATE != adStateClosed
            IF aWAData[ WA_CONNECTION ] :STATE != adStateOpen
               aWAData[ WA_CONNECTION ]:Cancel()
            ELSE
               aWAData[ WA_CONNECTION ]:Close()
            ENDIF
         ENDIF
      ENDIF
   RECOVER
   END SEQUENCE

   RETURN UR_SUPER_CLOSE( nWA )


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_GETVALUE()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_GETVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL rs      := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   IF aWAData[ WA_EOF ] .OR. rs:EOF .OR. rs:BOF
      xValue := NIL
      IF ADO_GETFIELDTYPE( rs:Fields( nField - 1 ) :Type ) == HB_FT_STRING
         xValue := Space( rs:Fields( nField - 1 ) :DefinedSize )
      ENDIF
   ELSE
      xValue := rs:Fields( nField - 1 ) :Value

      IF ADO_GETFIELDTYPE( rs:Fields( nField - 1 ) :Type ) == HB_FT_STRING
         IF ValType( xValue ) == "U"
            xValue := Space( rs:Fields( nField - 1 ) :DefinedSize )
         ELSE
            xValue := PadR( xValue, rs:Fields( nField - 1 ) :DefinedSize )
         ENDIF
      ELSEIF ADO_GETFIELDTYPE( rs:Fields( nField - 1 ) :Type ) == HB_FT_DATE
         /* Null values */
         IF ValType( xValue ) == "U"
            xValue := hb_SToD()
         ENDIF
      ELSEIF ADO_GETFIELDTYPE( rs:Fields( nField - 1 ) :Type ) == HB_FT_TIMESTAMP
         /* Null values */
         IF ValType( xValue ) == "U"
            xValue := hb_SToD()
         ENDIF
      ENDIF
   ENDIF

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_GOTO()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_GOTO( nWA, nRecord )

   LOCAL nRecNo
   LOCAL rs     := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   IF rs:RecordCount > 0
      rs:MoveFirst()
      rs:Move( nRecord - 1, 0 )
   ENDIF
   ADO_RECID( nWA, @nRecNo )

   RETURN iif( nRecord == nRecNo, HB_SUCCESS, HB_FAILURE )


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_GOTOID()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_GOTOID( nWA, nRecord )

   LOCAL nRecNo
   LOCAL rs     := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   IF rs:RecordCount > 0
      rs:MoveFirst()
      rs:Move( nRecord - 1, 0 )
   ENDIF
   ADO_RECID( nWA, @nRecNo )

   RETURN iif( nRecord == nRecNo, HB_SUCCESS, HB_FAILURE )


STATIC FUNCTION ADO_GOTOP( nWA )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF oRecordSet:RecordCount() != 0
      oRecordSet:MoveFirst()
      aWAData[ WA_BOF ] := .F.
      aWAData[ WA_EOF ] := .F.

      // Se parou num registro bloqueado pelo filtro, avança pro próximo válido
      IF aWAData[ WA_FILTERBLOCK ] != NIL .AND. !Eval( aWAData[ WA_FILTERBLOCK ] )
         ADO_SKIPRAW( nWA, 1 )
      ENDIF
   ELSE
      aWAData[ WA_BOF ] := .T.
      aWAData[ WA_EOF ] := .T.
   ENDIF

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_GOBOTTOM( nWA )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF oRecordSet:RecordCount() != 0
      oRecordSet:MoveLast()
      aWAData[ WA_BOF ] := .F.
      aWAData[ WA_EOF ] := .F.

      // Se parou num registro bloqueado pelo filtro, recua pro anterior válido
      IF aWAData[ WA_FILTERBLOCK ] != NIL .AND. !Eval( aWAData[ WA_FILTERBLOCK ] )
         ADO_SKIPRAW( nWA, -1 )
      ENDIF
   ELSE
      aWAData[ WA_BOF ] := .T.
      aWAData[ WA_EOF ] := .T.
   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_SKIPRAW( nWA, nToSkip )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult    := HB_SUCCESS
   LOCAL bFilter    := aWAData[ WA_FILTERBLOCK ]
   LOCAL nSkipped   := 0
   LOCAL nDirection := iif( nToSkip > 0, 1, -1 )

   IF !Empty( aWAData[ WA_PENDINGREL ] )
      IF ADO_FORCEREL( nWA ) != HB_SUCCESS
         RETURN HB_FAILURE
      ENDIF
   ENDIF

   IF nToSkip != 0
      IF aWAData[ WA_EOF ] .AND. nToSkip > 0
         RETURN HB_SUCCESS
      ENDIF

      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         IF aWAData[ WA_CONNECTION ]:STATE != 0 // adStateClosed
            DO WHILE nSkipped != nToSkip
               oRecordSet:Move( nDirection )

               IF oRecordSet:EOF
                  aWAData[ WA_EOF ] := .T.
                  aWAData[ WA_BOF ] := .F.
                  EXIT
               ENDIF
               IF oRecordSet:BOF
                  aWAData[ WA_BOF ] := .T.
                  aWAData[ WA_EOF ] := .F.
                  EXIT
               ENDIF

               // Testa o filtro Harbour nativo (passa direto se não houver filtro)
               IF bFilter == NIL .OR. Eval( bFilter )
                  nSkipped += nDirection
               ENDIF
            ENDDO
            aWAData[ WA_EOF ] := oRecordSet:EOF
            aWAData[ WA_BOF ] := oRecordSet:BOF
         ELSE
            nResult := HB_FAILURE
         ENDIF
      RECOVER
         nResult := HB_FAILURE
      END SEQUENCE
   ENDIF

   RETURN nResult


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_BOF()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_BOF( nWA, lBof )

   LOCAL aWAData := USRRDD_AREADATA( nWA )

   lBof := aWAData[ WA_BOF ]

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_EOF()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_EOF( nWA, lEof )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult    := HB_SUCCESS

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      IF USRRDD_AREADATA( nWA )[ WA_CONNECTION ] :STATE != adStateClosed
         lEof := ( oRecordSet:AbsolutePosition == -3 )
      ENDIF
   RECOVER
      nResult := HB_FAILURE
   END SEQUENCE

   RETURN nResult


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_DELETED()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_DELETED( nWA, lDeleted )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      IF oRecordSet:STATUS == adRecDeleted
         lDeleted := .T.
      ELSE
         lDeleted := .F.
      ENDIF
   RECOVER
      lDeleted := .F.
   END SEQUENCE

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_DELETE()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_DELETE( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   oRecordSet:Delete()

   ADO_SKIPRAW( nWA, 1 )

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RECNO()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RECNO( nWA, nRecNo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult    := HB_SUCCESS

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      IF USRRDD_AREADATA( nWA )[ WA_CONNECTION ] :STATE != adStateClosed
         nRecno := iif( oRecordSet:AbsolutePosition == -3, oRecordSet:RecordCount() + 1, oRecordSet:AbsolutePosition )
      ELSE
         nRecno  := 0
         nResult := HB_FAILURE
      ENDIF
   RECOVER
      nRecNo  := 0
      nResult := HB_FAILURE
   END SEQUENCE

   RETURN nResult


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RECID()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RECID( nWA, nRecNo )


   RETURN ADO_RECNO( nWA, @nRecNo )


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RECCOUNT()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RECCOUNT( nWA, nRecords )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   nRecords := oRecordSet:RecordCount()

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_PUTVALUE()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_PUTVALUE( nWA, nField, xValue )

   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF !aWAData[ WA_EOF ] .AND. !( oRecordSet:Fields( nField - 1 ) :Value == xValue )
      oRecordSet:Fields( nField - 1 ) :Value := xValue
       // ADO_PUTVALUE agora apenas armazena em cache (modo Batch)
      // O salvamento final (Update) ocorrerá no ADO_UNLOCK ou ADO_FLUSH
      //BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      //   oRecordSet:Update()
      //RECOVER
      //END SEQUENCE
   ENDIF

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_APPEND()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_APPEND( nWA, lUnLockAll )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   HB_SYMBOL_UNUSED( lUnLockAll )

   oRecordSet:AddNew()

   // O Update() foi removido daqui para permitir inserções em Batch super-rápidas
   //BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
   //   oRecordSet:Update()
   //RECOVER
   //END SEQUENCE

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_FLUSH()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_FLUSH( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:Update()
   RECOVER
   END SEQUENCE

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_ORDINFO()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_ORDINFO( nWA, nIndex, aOrderInfo )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult    := HB_SUCCESS

   DO CASE
   CASE nIndex == DBOI_EXPRESSION .OR. nIndex == DBOI_NAME
      IF !Empty( aWAData[ WA_CATALOG ] ) .AND. !Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
         aOrderInfo[ UR_ORI_TAG ] <= Len( aWAData[ WA_CATALOG ] )
         aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_CATALOG ][ aOrderInfo[ UR_ORI_TAG ] ]
      ELSE
         aOrderInfo[ UR_ORI_RESULT ] := ""
      ENDIF
   CASE nIndex == DBOI_NUMBER
      aOrderInfo[ UR_ORI_RESULT ] := aOrderInfo[ UR_ORI_TAG ]
   CASE nIndex == DBOI_BAGNAME .OR. nIndex == DBOI_BAGEXT
      aOrderInfo[ UR_ORI_RESULT ] := ""
   CASE nIndex == DBOI_ORDERCOUNT
      aOrderInfo[ UR_ORI_RESULT ] := iif( !Empty( aWAData[ WA_CATALOG ] ), Len( aWAData[ WA_CATALOG ] ), 0 )
   CASE nIndex == DBOI_FILEHANDLE
      aOrderInfo[ UR_ORI_RESULT ] := 0
   CASE nIndex == DBOI_ISCOND .OR. nIndex == DBOI_ISDESC .OR. nIndex == DBOI_UNIQUE
      aOrderInfo[ UR_ORI_RESULT ] := .F.
   CASE nIndex == DBOI_POSITION .OR. nIndex == DBOI_RECNO
      IF aWAData[ WA_CONNECTION ]:STATE != 0 // adStateClosed
         ADO_RECID( nWA, @aOrderInfo[ UR_ORI_RESULT ] )
      ELSE
         aOrderInfo[ UR_ORI_RESULT ] := 0
         nResult := HB_FAILURE
      ENDIF
   CASE nIndex == DBOI_KEYCOUNT
      IF aWAData[ WA_CONNECTION ]:STATE != 0
         aOrderInfo[ UR_ORI_RESULT ] := oRecordSet:RecordCount
      ELSE
         aOrderInfo[ UR_ORI_RESULT ] := 0
         nResult := HB_FAILURE
      ENDIF
   ENDCASE
   RETURN nResult

// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RECINFO()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RECINFO( nWA, nRecord, nInfoType, uInfo )

   LOCAL nResult := HB_SUCCESS

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   HB_SYMBOL_UNUSED( nWA )

#ifdef UR_DBRI_DELETED
   DO CASE
   CASE nInfoType == UR_DBRI_DELETED
      uInfo := .F.
   CASE nInfoType == UR_DBRI_LOCKED
      uInfo := .T.
   CASE nInfoType == UR_DBRI_RECSIZE
   CASE nInfoType == UR_DBRI_RECNO
      nResult := ADO_RECID( nWA, @nRecord )
   CASE nInfoType == UR_DBRI_UPDATED
      uInfo := .F.
   CASE nInfoType == UR_DBRI_ENCRYPTED
      uInfo := .F.
   CASE nInfoType == UR_DBRI_RAWRECORD
      uInfo := ""
   CASE nInfoType == UR_DBRI_RAWMEMOS
      uInfo := ""
   CASE nInfoType == UR_DBRI_RAWDATA
      nResult := ADO_GOTO( nWA, nRecord )
      uInfo   := ""
   ENDCASE
#else
   HB_SYMBOL_UNUSED( nRecord )
   HB_SYMBOL_UNUSED( nInfoType )
   HB_SYMBOL_UNUSED( uInfo )
#endif

   RETURN nResult


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_FIELDNAME()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_FIELDNAME( nWA, nField, cFieldName )

   LOCAL nResult    := HB_SUCCESS
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      cFieldName := oRecordSet:Fields( nField - 1 ) :Name
   RECOVER
      cFieldName := ""
      nResult    := HB_FAILURE
   END SEQUENCE

   RETURN nResult


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_FIELDINFO()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_FIELDINFO( nWA, nField, nInfoType, uInfo )

   LOCAL nType, nLen
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   DO CASE
   CASE nInfoType == DBS_NAME
      uInfo := oRecordSet:Fields( nField - 1 ) :Name

   CASE nInfoType == DBS_TYPE
      nType := ADO_GETFIELDTYPE( oRecordSet:Fields( nField - 1 ) :Type )
      DO CASE
      CASE nType == HB_FT_STRING
         uInfo := "C"
      CASE nType == HB_FT_LOGICAL
         uInfo := "L"
      CASE nType == HB_FT_MEMO
         uInfo := "M"
      CASE nType == HB_FT_OLE
         uInfo := "G"
#ifdef HB_FT_PICTURE
      CASE nType == HB_FT_PICTURE
         uInfo := "P"
#endif
      CASE nType == HB_FT_ANY
         uInfo := "V"
      CASE nType == HB_FT_DATE
         uInfo := "D"
#ifdef HB_FT_DATETIME
      CASE nType == HB_FT_DATETIME
         uInfo := "T"
#endif
      CASE nType == HB_FT_TIMESTAMP
         uInfo := "@"
      CASE nType == HB_FT_LONG
         uInfo := "N"
      CASE nType == HB_FT_INTEGER
         uInfo := "I"
      CASE nType == HB_FT_DOUBLE
         uInfo := "B"
      OTHERWISE
         uInfo := "U"
      ENDCASE

  CASE nInfoType == DBS_LEN
      ADO_FIELDINFO( nWA, nField, DBS_TYPE, @nType )
      uInfo := ADO_GETFIELDSIZE( oRecordSet:Fields( nField - 1 ), ADO_GETFIELDTYPE( oRecordSet:Fields( nField - 1 ) :Type ) )

   CASE nInfoType == DBS_DEC
      ADO_FIELDINFO( nWA, nField, DBS_TYPE, @nType )
      uInfo := ADO_GETFIELDDEC( oRecordSet:Fields( nField - 1 ), ADO_GETFIELDTYPE( oRecordSet:Fields( nField - 1 ) :Type ) )

#ifdef DBS_FLAG
   CASE nInfoType == DBS_FLAG
      uInfo := 0
#endif
#ifdef DBS_STEP
   CASE nInfoType == DBS_STEP
      uInfo := 0
#endif
   OTHERWISE
      RETURN HB_FAILURE
   ENDCASE

   RETURN HB_SUCCESS



STATIC FUNCTION ADO_ORDLSTFOCUS( nWA, aOrderInfo )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL cIndexName := ""

   IF aOrderInfo[ UR_ORI_TAG ] > 0 .AND. !Empty( aWAData[ WA_CATALOG ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= Len( aWAData[ WA_CATALOG ] )
      cIndexName := aWAData[ WA_CATALOG ][ aOrderInfo[ UR_ORI_TAG ] ]
   ENDIF

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:Index := cIndexName
      aOrderInfo[ UR_ORI_RESULT ] := aOrderInfo[ UR_ORI_TAG ]
   RECOVER
      aOrderInfo[ UR_ORI_RESULT ] := 0
      RETURN HB_FAILURE
   END SEQUENCE

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_PACK()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_PACK( nWA )

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   HB_SYMBOL_UNUSED( nWA )

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RAWLOCK()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RAWLOCK( nWA, nAction, nRecNo )

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

/* TODO */

   HB_SYMBOL_UNUSED( nRecNo )
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( nAction )

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_LOCK( nWA, aLockInfo )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nRecNo
   LOCAL lLocked    := .F.

   ADO_RECID( nWA, @nRecNo )

   IF !aWAData[ WA_EOF ] .AND. !aWAData[ WA_BOF ]
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         // O "Toque" (Touch): Força o SGBD a conceder um Lock Pessimista
         oRecordSet:Fields( 0 ):Value := oRecordSet:Fields( 0 ):Value
         lLocked := .T.
      RECOVER
         // Se cair aqui, outro usuário já está editando/travando este registro
         lLocked := .F.
      END SEQUENCE
   ELSE
      lLocked := .F.
   ENDIF

   aLockInfo[ UR_LI_METHOD ] := iif( aLockInfo[ UR_LI_METHOD ] == NIL, DBLM_MULTIPLE, aLockInfo[ UR_LI_METHOD ] )
   aLockInfo[ UR_LI_RECORD ] := nRecNo
   aLockInfo[ UR_LI_RESULT ] := lLocked

   RETURN HB_SUCCESS
   
   
 STATIC FUNCTION ADO_UNLOCK( nWA, xRecID )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   HB_SYMBOL_UNUSED( xRecId )

   IF !aWAData[ WA_EOF ] .AND. !aWAData[ WA_BOF ]
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         IF oRecordSet:EditMode != 0
            oRecordSet:Update()
         ENDIF
      RECOVER
         oRecordSet:CancelUpdate()
      END SEQUENCE
   ENDIF

   RETURN HB_SUCCESS  

STATIC FUNCTION ADO_SETFILTER( nWA, aFilterInfo )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cFilter := aFilterInfo[ UR_FRI_CEXPR ]

   IF !Empty( cFilter )
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         aWAData[ WA_FILTERBLOCK ] := &( "{|| " + cFilter + "}" )
      RECOVER
         aWAData[ WA_FILTERBLOCK ] := NIL
      END SEQUENCE
   ELSE
      aWAData[ WA_FILTERBLOCK ] := NIL
   ENDIF

   // Posiciona no primeiro registro válido que atende ao filtro
   ADO_GOTOP( nWA ) 

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_CLEARFILTER( nWA )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   
   aWAData[ WA_FILTERBLOCK ] := NIL
   ADO_GOTOP( nWA )

   RETURN HB_SUCCESS

// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_ZAP()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_ZAP( nWA )

   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF aWAData[ WA_CONNECTION ] != NIL .AND. aWAData[ WA_TABLENAME ] != NIL
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         aWAData[ WA_CONNECTION ]:Execute( "TRUNCATE TABLE " + aWAData[ WA_TABLENAME ] )
      RECOVER
         aWAData[ WA_CONNECTION ]:Execute( "DELETE * FROM " + aWAData[ WA_TABLENAME ] )
      END SEQUENCE
      oRecordSet:Requery()
   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_SETLOCATE( nWA, aScopeInfo )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   
   // Apenas guardamos o ScopeInfo sem usar SQLTranslate.
   // A avaliação será feita nativamente pela VM do Harbour.
   aWAData[ WA_SCOPEINFO ] := aScopeInfo

   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_LOCATE( nWA, lContinue )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL bFor, lFound := .F.
   LOCAL cForExpr   := aWAData[ WA_LOCATEFOR ]

   IF Empty( cForExpr )
      RETURN HB_FAILURE
   ENDIF

   // Compila a expressão em Harbour puro e seguro
   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      bFor := &( "{|| " + cForExpr + "}" )
   RECOVER
      RETURN HB_FAILURE
   END SEQUENCE

   // Prepara o ponteiro (Início ou Próximo Registro)
   IF lContinue
      ADO_SKIPRAW( nWA, 1 )
   ELSE
      ADO_GOTOP( nWA )
   ENDIF

   // Loop de varredura nativa super-rápida (em cache local do ADO)
   DO WHILE !aWAData[ WA_EOF ]
      IF Eval( bFor )
         lFound := .T.
         EXIT
      ENDIF
      ADO_SKIPRAW( nWA, 1 )
   ENDDO

   aWAData[ WA_FOUND ] := lFound

   RETURN HB_SUCCESS   

// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_CLEARREL()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_CLEARREL( nWA )

   /*
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nKeys   := 0, cKeyName

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      nKeys := aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ) :Keys:Count
   RECOVER
   END SEQUENCE

   IF nKeys > 0
      cKeyName := aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys( nKeys - 1 ) :Name
      IF !( Upper( cKeyName ) == "PRIMARYKEY" )
         aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys:Delete( cKeyName )
      ENDIF

   ENDIF
   */

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RELAREA()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RELAREA( nWA, nRelNo, nRelArea )

   /*
   LOCAL aWAData := USRRDD_AREADATA( nWA )

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      IF nRelNo <= aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys:Count()
         nRelArea := SELECT ( aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys( nRelNo - 1 ) :RelatedTable )
      ENDIF
   RECOVER
      nRelArea := 0
   END SEQUENCE
   */ 
   nRelArea := 0 
   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RELTEXT()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RELTEXT( nWA, nRelNo, cExpr )

   /*
   LOCAL aWAData := USRRDD_AREADATA( nWA )

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      IF nRelNo <= aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys:Count()
         cExpr := aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys( nRelNo - 1 ):Columns( 0 ) :RelatedColumn
      ENDIF
   RECOVER
      cExpr := ""
   END SEQUENCE
   */
   cExpr := ""
   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_SETREL()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_SETREL( nWA, aRelInfo )

   /*
   LOCAL aWAData  := USRRDD_AREADATA( nWA )
   LOCAL cParent  := Alias( aRelInfo[ UR_RI_PARENT ] )
   LOCAL cChild   := Alias( aRelInfo[ UR_RI_CHILD ] )
   LOCAL cKeyName := cParent + "_" + cChild

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      aWAData[ WA_CATALOG ]:Tables( aWAData[ WA_TABLENAME ] ):Keys:Append( cKeyName, adKeyForeign, ;
         aRelInfo[ UR_RI_CEXPR ], cChild, aRelInfo[ UR_RI_CEXPR ] )
   RECOVER
      /* raise error for can't create relation */
//   END SEQUENCE
   */

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_FORCEREL()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_FORCEREL( nWA )

   LOCAL aPendingRel
   LOCAL aWAData     := USRRDD_AREADATA( nWA )

   IF !Empty( aWAData[ WA_PENDINGREL ] )

      aPendingRel              := aWAData[ WA_PENDINGREL ]
      aWAData[ WA_PENDINGREL ] := NIL

      RETURN ADO_RELEVAL( nWA, aPendingRel )
   ENDIF

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_RELEVAL()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_RELEVAL( nWA, aRelInfo )

   LOCAL aInfo, nReturn, nOrder, uResult

   nReturn := ADO_EVALBLOCK( aRelInfo[ UR_RI_PARENT ], aRelInfo[ UR_RI_BEXPR ], @uResult )

   IF nReturn == HB_SUCCESS
   /*
       *  Check the current order
       */
      aInfo   := Array( UR_ORI_SIZE )
      nReturn := ADO_ORDINFO( nWA, DBOI_NUMBER, @aInfo )

      IF nReturn == HB_SUCCESS
         nOrder := aInfo[ UR_ORI_RESULT ]
         IF nOrder != 0
            IF aRelInfo[ UR_RI_SCOPED ]
               aInfo[ UR_ORI_NEWVAL ] := uResult
               nReturn                := ADO_ORDINFO( nWA, DBOI_SCOPETOP, @aInfo )
               IF nReturn == HB_SUCCESS
                  nReturn := ADO_ORDINFO( nWA, DBOI_SCOPEBOTTOM, @aInfo )
               ENDIF
            ENDIF
            IF nReturn == HB_SUCCESS
               nReturn := ADO_SEEK( nWA, .F., uResult, .F. )
            ENDIF
         ELSE
            nReturn := ADO_GOTO( nWA, uResult )
         ENDIF
      ENDIF
   ENDIF

   RETURN nReturn


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_ORDLSTADD()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_ORDLSTADD( nWA, aOrderInfo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:INDEX := aOrderInfo[ UR_ORI_BAG ]
   RECOVER
   END SEQUENCE

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_ORDLSTCLEAR()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_ORDLSTCLEAR( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:INDEX := ""
   RECOVER
   END SEQUENCE

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_ORDCREATE( nWA, aOrderCreateInfo )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cIndexName, cKeyExpr, cSql, oError

   cIndexName := iif( !Empty( aOrderCreateInfo[ UR_ORCR_TAGNAME ] ), aOrderCreateInfo[ UR_ORCR_TAGNAME ], aOrderCreateInfo[ UR_ORCR_CKEY ] )
   cKeyExpr   := aOrderCreateInfo[ UR_ORCR_CKEY ]

   cSql := "CREATE " + iif( aOrderCreateInfo[ UR_ORCR_UNIQUE ], "UNIQUE ", "" ) + ;
           "INDEX " + cIndexName + " ON " + aWAData[ WA_TABLENAME ] + " (" + cKeyExpr + ")"

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      aWAData[ WA_CONNECTION ]:Execute( cSql )
      IF AScan( aWAData[ WA_CATALOG ], {|x| Upper(x) == Upper(cIndexName)} ) == 0
         AAdd( aWAData[ WA_CATALOG ], cIndexName )
      ENDIF
   RECOVER USING oError
      oError:GenCode     := EG_CREATE
      oError:SubCode     := 1004
      oError:Description := "Erro SQL ao criar Indice: " + cIndexName
      UR_SUPER_ERROR( nWA, oError )
   END SEQUENCE
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDDESTROY( nWA, aOrderInfo )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cIndexName := aOrderInfo[ UR_ORI_TAG ]
   LOCAL cSql, nPos

   // Tratamento de dialeto para exclusão de índice
   IF t_cEngine == "PGSQL" .OR. t_cEngine == "POSTGRESQL" .OR. t_cEngine == "SQLITE" .OR. t_cEngine == "MYSQL" .OR. t_cEngine == "MARIADB"
      cSql := "DROP INDEX " + cIndexName + " ON " + aWAData[ WA_TABLENAME ]
   ELSE
      // SQL Server, Access, Oracle...
      cSql := "DROP INDEX " + aWAData[ WA_TABLENAME ] + "." + cIndexName
   ENDIF

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      aWAData[ WA_CONNECTION ]:Execute( cSql )
      nPos := AScan( aWAData[ WA_CATALOG ], {|x| Upper(x) == Upper(cIndexName)} )
      IF nPos > 0
         ADel( aWAData[ WA_CATALOG ], nPos )
         ASize( aWAData[ WA_CATALOG ], Len( aWAData[ WA_CATALOG ] ) - 1 )
      ENDIF
   RECOVER
   END SEQUENCE
   RETURN HB_SUCCESS

// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_EVALBLOCK()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_EVALBLOCK( nArea, bBlock, uResult )

   LOCAL nCurrArea

   nCurrArea := SELECT ()
   IF nCurrArea != nArea
      dbSelectArea( nArea )
   ELSE
      nCurrArea := 0
   ENDIF

   uResult := Eval( bBlock )

   IF nCurrArea > 0
      dbSelectArea( nCurrArea )
   ENDIF

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_EXISTS()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_EXISTS( nRdd, cTable, cIndex, ulConnect )

// LOCAL n
   LOCAL lRet   := HB_FAILURE
   LOCAL aRData := USRRDD_RDDDATA( nRDD )

   HB_SYMBOL_UNUSED( ulConnect )

   IF !Empty( cTable ) .AND. !Empty( aRData[ WA_CATALOG ] )
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         // n := aRData[ WA_CATALOG ]:Tables( cTable )
         lRet := HB_SUCCESS
      RECOVER
      END SEQUENCE
      IF !Empty( cIndex )
         BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
            // n := aRData[ WA_CATALOG ]:Tables( cTable ):Indexes( cIndex )
            lRet := HB_SUCCESS
         RECOVER
         END SEQUENCE
      ENDIF
   ENDIF

   RETURN lRet


// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_DROP()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_DROP( nRdd, cTable, cIndex, ulConnect )

// LOCAL n
   LOCAL lRet   := HB_FAILURE
   LOCAL aRData := USRRDD_RDDDATA( nRDD )

   HB_SYMBOL_UNUSED( ulConnect )

   IF !Empty( cTable ) .AND. !Empty( aRData[ WA_CATALOG ] )
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         // n := aRData[ WA_CATALOG ]:Tables:Delete( cTable )
         lRet := HB_SUCCESS
      RECOVER
      END SEQUENCE
      IF !Empty( cIndex )
         BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
            // n := aRData[ WA_CATALOG ]:Tables( cTable ):Indexes:Delete( cIndex )
            lRet := HB_SUCCESS
         RECOVER
         END SEQUENCE
      ENDIF
   ENDIF

   RETURN lRet


STATIC FUNCTION ADO_SEEK( nWA, lSoftSeek, xKey, lFindLast )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL cFieldName := ""
   LOCAL cFilter    := ""
   LOCAL nType

   HB_SYMBOL_UNUSED( lSoftSeek )
   HB_SYMBOL_UNUSED( lFindLast )

   // 1. Identifica a coluna a pesquisar baseada no Índice ativo
   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      cFieldName := oRecordSet:Index
   RECOVER
      cFieldName := ""
   END SEQUENCE

   // Fallback: se não houver índice setado, pesquisa na primeira coluna da tabela
   IF Empty( cFieldName )
      cFieldName := oRecordSet:Fields( 0 ):Name
   ENDIF

   // 2. Trata a formatação correta baseada no tipo de dado pesquisado
   nType := ValType( xKey )
   DO CASE
   CASE nType == "C"
      // Dobra aspas simples na string para evitar quebra (SQL Injection Acidental)
      cFilter := cFieldName + " = '" + StrTran( xKey, "'", "''" ) + "'"
   CASE nType == "N"
      cFilter := cFieldName + " = " + hb_ntos( xKey )
   CASE nType == "D"
      // Converte data Harbour para formato universal ANSI 'YYYY-MM-DD' reconhecido pelo ADO
      cFilter := cFieldName + " = '" + Transform( DtoS( xKey ), "@R 9999-99-99" ) + "'"
   CASE nType == "L"
      cFilter := cFieldName + " = " + iif( xKey, "True", "False" )
   OTHERWISE
      cFilter := cFieldName + " = '" + hb_ValToStr( xKey ) + "'"
   ENDCASE

   // 3. Executa a busca isolada para evitar travar se o cFieldName for uma expressão complexa
   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRecordSet:MoveFirst()
      oRecordSet:Find( cFilter )
   RECOVER
      aWAData[ WA_FOUND ] := .F.
      RETURN HB_FAILURE
   END SEQUENCE

   IF oRecordSet:EOF
      aWAData[ WA_FOUND ] := .F.
      RETURN HB_FAILURE
   ENDIF

   aWAData[ WA_FOUND ] := .T.
   RETURN HB_SUCCESS   

// +--------------------------------------------------------------------
// +
// +
// +
// +    Static Function ADO_FOUND()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
STATIC FUNCTION ADO_FOUND( nWA, lFound )

   LOCAL aWAData := USRRDD_AREADATA( nWA )

   lFound := aWAData[ WA_FOUND ]

   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +
// +
// +    Function RDDADOX_GETFUNCTABLE()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
FUNCTION RDDADOX_GETFUNCTABLE( pFuncCount, pFuncTable, pSuperTable, nRddID )

   LOCAL aADOFunc[ UR_METHODCOUNT ]

   aADOFunc[ UR_INIT ]         := @ADO_INIT()
   aADOFunc[ UR_NEW ]          := @ADO_NEW()
   aADOFunc[ UR_CREATE ]       := @ADO_CREATE()
   aADOFunc[ UR_CREATEFIELDS ] := @ADO_CREATEFIELDS()
   aADOFunc[ UR_OPEN ]         := @ADO_OPEN()
   aADOFunc[ UR_CLOSE ]        := @ADO_CLOSE()
   aADOFunc[ UR_BOF ]          := @ADO_BOF()
   aADOFunc[ UR_EOF ]          := @ADO_EOF()
   aADOFunc[ UR_DELETED ]      := @ADO_DELETED()
   aADOFunc[ UR_SKIPRAW ]      := @ADO_SKIPRAW()
   aADOFunc[ UR_GOTO ]         := @ADO_GOTO()
   aADOFunc[ UR_GOTOID ]       := @ADO_GOTOID()
   aADOFunc[ UR_GOTOP ]        := @ADO_GOTOP()
   aADOFunc[ UR_GOBOTTOM ]     := @ADO_GOBOTTOM()
   aADOFunc[ UR_RECNO ]        := @ADO_RECNO()
   aADOFunc[ UR_RECID ]        := @ADO_RECID()
   aADOFunc[ UR_RECCOUNT ]     := @ADO_RECCOUNT()
   aADOFunc[ UR_GETVALUE ]     := @ADO_GETVALUE()
   aADOFunc[ UR_PUTVALUE ]     := @ADO_PUTVALUE()
   aADOFunc[ UR_DELETE ]       := @ADO_DELETE()
   aADOFunc[ UR_APPEND ]       := @ADO_APPEND()
   aADOFunc[ UR_FLUSH ]        := @ADO_FLUSH()
   aADOFunc[ UR_ORDINFO ]      := @ADO_ORDINFO()
   aADOFunc[ UR_RECINFO ]      := @ADO_RECINFO()
   aADOFunc[ UR_FIELDINFO ]    := @ADO_FIELDINFO()
   aADOFunc[ UR_FIELDNAME ]    := @ADO_FIELDNAME()
   aADOFunc[ UR_ORDLSTFOCUS ]  := @ADO_ORDLSTFOCUS()
   aADOFunc[ UR_PACK ]         := @ADO_PACK()
   aADOFunc[ UR_RAWLOCK ]      := @ADO_RAWLOCK()
   aADOFunc[ UR_LOCK ]         := @ADO_LOCK()
   aADOFunc[ UR_UNLOCK ]       := @ADO_UNLOCK()
   aADOFunc[ UR_SETFILTER ]    := @ADO_SETFILTER()
   aADOFunc[ UR_CLEARFILTER ]  := @ADO_CLEARFILTER()
   aADOFunc[ UR_ZAP ]          := @ADO_ZAP()
   aADOFunc[ UR_SETLOCATE ]    := @ADO_SETLOCATE()
   aADOFunc[ UR_LOCATE ]       := @ADO_LOCATE()
   aADOFunc[ UR_FOUND ]        := @ADO_FOUND()
   aADOFunc[ UR_FORCEREL ]     := @ADO_FORCEREL()
   aADOFunc[ UR_RELEVAL ]      := @ADO_RELEVAL()
   aADOFunc[ UR_CLEARREL ]     := @ADO_CLEARREL()
   aADOFunc[ UR_RELAREA ]      := @ADO_RELAREA()
   aADOFunc[ UR_RELTEXT ]      := @ADO_RELTEXT()
   aADOFunc[ UR_SETREL ]       := @ADO_SETREL()
   aADOFunc[ UR_ORDCREATE ]    := @ADO_ORDCREATE()
   aADOFunc[ UR_ORDDESTROY ]   := @ADO_ORDDESTROY()
   aADOFunc[ UR_ORDLSTADD ]    := @ADO_ORDLSTADD()
   aADOFunc[ UR_ORDLSTCLEAR ]  := @ADO_ORDLSTCLEAR()
   aADOFunc[ UR_EVALBLOCK ]    := @ADO_EVALBLOCK()
   aADOFunc[ UR_SEEK ]         := @ADO_SEEK()
   aADOFunc[ UR_EXISTS ]       := @ADO_EXISTS()
   aADOFunc[ UR_DROP ]         := @ADO_DROP()
   aADOFunc[ UR_TRANSBEGIN ]    := @ADO_TRANSBEGIN()
   aADOFunc[ UR_TRANSCOMMIT ]   := @ADO_TRANSCOMMIT()
   aADOFunc[ UR_TRANSROLLBACK ] := @ADO_TRANSROLLBACK()
   aADOFunc[ UR_GETROW      ] := @ADO_GETROW()
   aADOFunc[ UR_GETROWBLANK ] := @ADO_GETROWBLANK()
   aADOFunc[ UR_PUTROW      ] := @ADO_PUTROW()

   RETURN USRRDD_GETFUNCTABLE( pFuncCount, pFuncTable, pSuperTable, nRddID, ;
      /* NO SUPER RDD */,aADOFunc)


// +--------------------------------------------------------------------
// +
// +
// +
// +    Init Procedure RDDADOX_INIT()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
INIT PROCEDURE RDDADOX_INIT()

   rddRegister( "RDDADOX", RDT_FULL )

   RETURN


// +--------------------------------------------------------------------
// +    Static Function ADO_GETFIELDSIZE() e ADO_GETFIELDDEC()
// +    Corrigidas para ler a precisão e escala nativa (NumericScale) do ADO
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_GETFIELDSIZE( oField, nDBFFieldType )
   LOCAL nDBFFieldSize := 0

   DO CASE
   CASE nDBFFieldType == HB_FT_STRING
      nDBFFieldSize := oField:DefinedSize
      // Proteção contra campos de texto estourados (> 1024 vira Memo DBF)
      IF nDBFFieldSize > 1024 .OR. nDBFFieldSize < 0
         nDBFFieldSize := 10
      ENDIF
   CASE nDBFFieldType == HB_FT_INTEGER .OR. nDBFFieldType == HB_FT_LONG .OR. nDBFFieldType == HB_FT_DOUBLE
      IF oField:Precision > 0
         nDBFFieldSize := oField:Precision
      ELSE
         nDBFFieldSize := iif( nDBFFieldType == HB_FT_DOUBLE, 15, 11 )
      ENDIF
   CASE nDBFFieldType == HB_FT_DATE .OR. nDBFFieldType == HB_FT_TIMESTAMP
      nDBFFieldSize := 8
   CASE nDBFFieldType == HB_FT_LOGICAL
      nDBFFieldSize := 1
   CASE nDBFFieldType == HB_FT_MEMO .OR. nDBFFieldType == HB_FT_OLE
      nDBFFieldSize := 10
   OTHERWISE
      nDBFFieldSize := oField:DefinedSize
   ENDCASE

   RETURN nDBFFieldSize

STATIC FUNCTION ADO_GETFIELDDEC( oField, nDBFFieldType )
   LOCAL nDec := 0
   IF nDBFFieldType == HB_FT_DOUBLE .OR. nDBFFieldType == HB_FT_LONG .OR. nDBFFieldType == HB_FT_INTEGER
      nDec := oField:NumericScale
      // O ADO retorna 255 quando o banco não define escala fixa
      IF nDec == 255 .OR. nDec < 0
         nDec := 0
      ENDIF
   ENDIF
   RETURN nDec

// +--------------------------------------------------------------------
// +    Static Function ADO_GETFIELDTYPE()
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_GETFIELDTYPE( nADOFieldType )
   LOCAL nDBFFieldType := HB_FT_STRING 

   DO CASE
   CASE nADOFieldType == adBoolean
      nDBFFieldType := HB_FT_LOGICAL
   CASE nADOFieldType == adDate .OR. nADOFieldType == adDBDate
      nDBFFieldType := HB_FT_DATE
   CASE nADOFieldType == adDBTime .OR. nADOFieldType == adDBTimeStamp .OR. nADOFieldType == adFileTime
      nDBFFieldType := HB_FT_TIMESTAMP
   CASE nADOFieldType == adTinyInt .OR. nADOFieldType == adSmallInt .OR. nADOFieldType == adInteger .OR. nADOFieldType == adBigInt .OR. nADOFieldType == adUnsignedTinyInt .OR. nADOFieldType == adUnsignedSmallInt .OR. nADOFieldType == adUnsignedInt .OR. nADOFieldType == adUnsignedBigInt
      nDBFFieldType := HB_FT_INTEGER
   CASE nADOFieldType == adSingle .OR. nADOFieldType == adDouble .OR. nADOFieldType == adCurrency .OR. nADOFieldType == adDecimal .OR. nADOFieldType == adNumeric .OR. nADOFieldType == adVarNumeric
      nDBFFieldType := HB_FT_DOUBLE
   CASE nADOFieldType == adBSTR .OR. nADOFieldType == adChar .OR. nADOFieldType == adVarChar .OR. nADOFieldType == adWChar .OR. nADOFieldType == adVarWChar .OR. nADOFieldType == adGUID
      nDBFFieldType := HB_FT_STRING
   CASE nADOFieldType == adLongVarChar .OR. nADOFieldType == adLongVarWChar .OR. nADOFieldType == adPropVariant
      nDBFFieldType := HB_FT_MEMO
   CASE nADOFieldType == adBinary .OR. nADOFieldType == adVarBinary .OR. nADOFieldType == adLongVarBinary
      nDBFFieldType := HB_FT_OLE
   OTHERWISE
      nDBFFieldType := HB_FT_ANY
   ENDCASE

   RETURN nDBFFieldType

// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetTable()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetTable( cTableName )

   t_cTableName := cTableName
   //t_cTableName := AllTrim( cTableName )

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetEngine()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetEngine( cEngine )

   t_cEngine := cEngine
   
   t_cEngine := Upper( AllTrim( cEngine ) )

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetServer()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetServer( cServer )

   t_cServer := cServer

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetUser()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetUser( cUser )

   t_cUserName := cUser

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetPassword()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetPassword( cPassword )

   t_cPassword := cPassword

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetQuery()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetQuery( cQuery )

   hb_default( @cQuery, "SELECT * FROM " )

   t_cQuery := cQuery

   RETURN


// +--------------------------------------------------------------------
// +
// +
// +
// +    Procedure hb_adoSetLocateFor()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
PROCEDURE hb_adoSetLocateFor( cLocateFor )

   USRRDD_AREADATA( Select() )[ WA_LOCATEFOR ] := cLocateFor

   RETURN


// +--------------------------------------------------------------------
// +
// +    Static Function SQLTranslate()
// +    Tradutor mestre de expressões Harbour para o Dialeto SQL correto
// +
// +--------------------------------------------------------------------
STATIC FUNCTION SQLTranslate( cExpr )

   // 1. Limpeza nativa de aspas da RDDADOX
   IF Left( cExpr, 1 ) == '"' .AND. Right( cExpr, 1 ) == '"'
      cExpr := SubStr( cExpr, 2, Len( cExpr ) - 2 )
   ENDIF
   cExpr := StrTran( cExpr, '""' )
   cExpr := StrTran( cExpr, '"', "'" )
   cExpr := StrTran( cExpr, "''", "'" )
   cExpr := StrTran( cExpr, "==", "=" )

   // 2. Aplica as regras de dialeto da sua biblioteca dbudialeto
   cExpr := ADO_DIALETO_CONDICIONAIS( cExpr, t_cEngine )
   cExpr := ADO_CONVERTER_EMPTY( cExpr, t_cEngine )
   cExpr := ADO_DIALETO_FUNCOES( cExpr, t_cEngine )

   RETURN cExpr

// +--------------------------------------------------------------------
// +    Rotinas de Apoio Autocontidas (Baseadas no dbudialeto.prg)
// +--------------------------------------------------------------------

STATIC FUNCTION ADO_DIALETO_CONDICIONAIS( cSQLCNV, cEngine )
   cSQLCNV := StrTran( cSQLCNV, ".and.", " AND " )
   cSQLCNV := StrTran( cSQLCNV, ".or.", " OR " )
   cSQLCNV := StrTran( cSQLCNV, ".AND.", " AND " )
   cSQLCNV := StrTran( cSQLCNV, ".OR.", " OR " )
   cSQLCNV := StrTran( cSQLCNV, ".NOT.", " NOT " )
   cSQLCNV := StrTran( cSQLCNV, ".not.", " NOT " )

   DO CASE
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cSQLCNV := StrTran( cSQLCNV, "!=", " <> " )
   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cSQLCNV := StrTran( cSQLCNV, "<>", " != " )
   ENDCASE
   RETURN cSQLCNV

STATIC FUNCTION ADO_DIALETO_FUNCOES( cSQLCNV, cEngine )
   DO CASE
   CASE cEngine == "SQLITE"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "CURRENT_DATE " )
      cSQLCNV := StrTran( cSQLCNV, "CHR(", "CHAR(" )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "TRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "LEN(", "LENGTH(" )
      cSQLCNV := StrTran( cSQLCNV, "CURRENTDATETIME", " current_timestamp " )
      cSQLCNV := StrTran( cSQLCNV, "REPL(", "printf('%.*c', " ) 
      cSQLCNV := StrTran( cSQLCNV, "SUBSTR(", "SUBSTR(" ) 
      cSQLCNV := StrTran( cSQLCNV, "DTOS(", "strftime('%Y%m%d', " )

   CASE cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "SYSDATE()" )
      cSQLCNV := StrTran( cSQLCNV, "CHR(", "CHAR(" )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "TRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "REPL(", "REPEAT(" )

   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "CURRENT_DATE " )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "TRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "LEN(", "LENGTH(" )
      cSQLCNV := StrTran( cSQLCNV, "DAY(", "EXTRACT('DAY' FROM " )
      cSQLCNV := StrTran( cSQLCNV, "MONTH(", "EXTRACT('MONTH' FROM " )
      cSQLCNV := StrTran( cSQLCNV, "YEAR(", "EXTRACT('YEAR' FROM " )
      cSQLCNV := StrTran( cSQLCNV, "REPL(", "REPEAT(" )
      cSQLCNV := StrTran( cSQLCNV, ".T.", "TRUE" )
      cSQLCNV := StrTran( cSQLCNV, ".F.", "FALSE" )
      cSQLCNV := StrTran( cSQLCNV, "'  /  /  '", "NULL" )
      cSQLCNV := StrTran( cSQLCNV, "'00/00/0000'", "NULL" )
      cSQLCNV := StrTran( cSQLCNV, "IIF(", "CASE WHEN " )
      cSQLCNV := StrTran( cSQLCNV, "NOW()", "CURRENT_TIMESTAMP" )
      cSQLCNV := StrTran( cSQLCNV, "IFNULL(", "COALESCE(" )
     
   CASE cEngine == "DUCKDB" .OR. cEngine == "DUCKLAKE"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "CURRENT_DATE" )
      cSQLCNV := StrTran( cSQLCNV, "LEN(", "LENGTH(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "TRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "CHR(", "CHR(" )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "SUBSTR(", "SUBSTRING(" )
      cSQLCNV := StrTran( cSQLCNV, "YEAR(", "EXTRACT('YEAR' FROM " )
      cSQLCNV := StrTran( cSQLCNV, "MONTH(", "EXTRACT('MONTH' FROM " )
      cSQLCNV := StrTran( cSQLCNV, "DAY(", "EXTRACT('DAY' FROM " )
      
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "GETDATE() " )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "TRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "REPL(", "REPLICATE(" )
      cSQLCNV := StrTran( cSQLCNV, "CHR(", "CHAR(" )
      cSQLCNV := StrTran( cSQLCNV, "SUBSTR(", "SUBSTRING(" )
      cSQLCNV := StrTran( cSQLCNV, "AT(", "CHARINDEX(" )

   CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
      cSQLCNV := StrTran( cSQLCNV, "TODAY()", "SYSDATE " )
      cSQLCNV := StrTran( cSQLCNV, "CHR(", "CHAR(" )
      cSQLCNV := StrTran( cSQLCNV, "ASC(", "ASCII(" )
      cSQLCNV := StrTran( cSQLCNV, "TRIM(", "RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "ALLTRIM(", "LTRIM(RTRIM(" )
      cSQLCNV := StrTran( cSQLCNV, "LEN(", "LENGTH(" )
      cSQLCNV := StrTran( cSQLCNV, "REPL(", "REPLICATE(" )

   CASE cEngine == "ACCESS" .OR. cEngine == "MDB" .OR. cEngine == "ACCDB" .OR. cEngine == "ACEOLEDB"
      cSQLCNV := StrTran( cSQLCNV, "CURRENTDATETIME", " now " )
   ENDCASE

   RETURN cSQLCNV

STATIC FUNCTION ADO_CONVERTER_EMPTY( cSQL, cEngine )
   LOCAL nPos, nInicio, nFim, cCampo, cSubst, lNot, nTamRemover
   
   DO WHILE (At("EMPTY(", Upper(cSQL)) > 0)
      nPos := At("EMPTY(", Upper(cSQL))
      lNot := ADO_DETECTAR_NEGACAO(cSQL, nPos)
      
      nInicio := At("(", SubStr(cSQL, nPos)) + nPos - 1
      nFim    := At(")", SubStr(cSQL, nInicio))
      
      IF nFim > 0
         nFim += nInicio - 1
         cCampo := AllTrim(SubStr(cSQL, nInicio + 1, nFim - nInicio - 1))
         
         IF lNot
            nTamRemover := IIF(At("NOT", Upper(SubStr(cSQL, Max(1, nPos-4), 4))) > 0, 3, 1)
            cSQL := Stuff(cSQL, nPos - nTamRemover, (nFim - (nPos - nTamRemover) + 1), "")
         ELSE
            cSQL := Stuff(cSQL, nPos, (nFim - nPos + 1), "")
         ENDIF
         
         cSubst := ADO_GERAR_FRAGMENTO_SQL(cCampo, lNot, cEngine)
         cSQL := SubStr(cSQL, 1, nPos - 1 - IIF(lNot, nTamRemover, 0)) + cSubst + SubStr(cSQL, nPos - IIF(lNot, nTamRemover, 0))
      ELSE
         EXIT 
      ENDIF
   ENDDO
   RETURN cSQL

STATIC FUNCTION ADO_DETECTAR_NEGACAO( cSQL, nPos )
   LOCAL cPrecedente := ""
   LOCAL lNot := .F.
   IF nPos > 5
      cPrecedente := AllTrim(Upper(SubStr(cSQL, nPos - 5, 5)))
   ELSE
      cPrecedente := AllTrim(Upper(SubStr(cSQL, 1, nPos - 1)))
   ENDIF
   IF Right(cPrecedente, 1) == "!" .OR. Right(cPrecedente, 3) == "NOT"
      lNot := .T.
   ENDIF
   RETURN lNot

STATIC FUNCTION ADO_GERAR_FRAGMENTO_SQL(cCampo, lNot, cEngine)
   LOCAL cRet := ""
   IF lNot
      DO CASE
         CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
            cRet := " ( " + cCampo + " IS NOT NULL ) "
         OTHERWISE 
            cRet := " ( " + cCampo + " IS NOT NULL AND " + cCampo + " <> '' ) "
      ENDCASE
   ELSE
      DO CASE
         CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
            cRet := " ( " + cCampo + " IS NULL ) "
         OTHERWISE 
            cRet := " ( " + cCampo + " IS NULL OR " + cCampo + " = '' ) "
      ENDCASE
   ENDIF
   RETURN cRet

// +--------------------------------------------------------------------
// +
// +
// +
// +    Function hb_RDDADOXGetConnection()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
FUNCTION hb_RDDADOXGetConnection( nWA )

   IF !HB_ISNUMERIC( nWA )
      nWA := SELECT ()
   ENDIF

   RETURN USRRDD_AREADATA( nWA )[ WA_CONNECTION ]


// +--------------------------------------------------------------------
// +
// +
// +
// +    Function hb_RDDADOXGetCatalog()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
FUNCTION hb_RDDADOXGetCatalog( nWA )

   IF !HB_ISNUMERIC( nWA )
      nWA := SELECT ()
   ENDIF

   RETURN USRRDD_AREADATA( nWA )[ WA_CATALOG ]


// +--------------------------------------------------------------------
// +
// +
// +
// +    Function hb_RDDADOXGetRecordSet()
// +
// +
// +
// +--------------------------------------------------------------------
// +
// +
// +
FUNCTION hb_RDDADOXGetRecordSet( nWA )

   LOCAL aWAData

   IF !HB_ISNUMERIC( nWA )
      nWA := SELECT ()
   ENDIF

   aWAData := USRRDD_AREADATA( nWA )

   RETURN iif( aWAData != NIL, aWAData[ WA_RECORDSET ], NIL )


// +--------------------------------------------------------------------
// +
// +    Novas Funções de Suporte a Transações Nativas (Dialeto)
// +
// +--------------------------------------------------------------------

STATIC FUNCTION ADO_DIALETO_BEGIN( cEngine )
   LOCAL cComando := "BEGIN TRANSACTION"
   DO CASE
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cComando := "BEGIN TRANSACTION"
   CASE cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
      cComando := "START TRANSACTION;"
   CASE cEngine == "FIREBIRD" .OR. cEngine == "FDB" .OR. cEngine == "GDB" .OR. cEngine == "IB"
      cComando := "SET TRANSACTION"
   CASE cEngine == "SQLITE"
      cComando := "BEGIN TRANSACTION;"
   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cComando := "BEGIN;"
   CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
      cComando := "SET TRANSACTION READ WRITE;"
   CASE cEngine == "DUCKDB" .OR. cEngine == "DUCKLAKE"
      cComando := "BEGIN TRANSACTION;"
   ENDCASE
   RETURN cComando

STATIC FUNCTION ADO_DIALETO_COMMIT( cEngine )
   LOCAL cComando := "COMMIT TRANSACTION"
   DO CASE
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cComando := "IF @@TRANCOUNT > 0 COMMIT"
   CASE cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
      cComando := "COMMIT;"
   CASE cEngine == "FIREBIRD" .OR. cEngine == "FDB" .OR. cEngine == "GDB" .OR. cEngine == "IB"
      cComando := "COMMIT"
   CASE cEngine == "SQLITE"
      cComando := "end transaction"
   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cComando := "COMMIT;"
   CASE cEngine == "DUCKDB" .OR. cEngine == "DUCKLAKE"
      cComando := "COMMIT;"
   ENDCASE
   RETURN cComando

STATIC FUNCTION ADO_DIALETO_ROLLBACK( cEngine )
   LOCAL cComando := "ROLLBACK TRANSACTION"
   DO CASE
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cComando := "IF @@TRANCOUNT > 0 ROLLBACK"
   CASE cEngine == "MYSQL" .OR. cEngine == "MYSQL64" .OR. cEngine == "MARIADB"
      cComando := "ROLLBACK;"
   CASE cEngine == "FIREBIRD" .OR. cEngine == "FDB" .OR. cEngine == "GDB" .OR. cEngine == "IB"
      cComando := "ROLLBACK"
   CASE cEngine == "SQLITE"
      cComando := "ROLLBACK;"
   CASE cEngine == "PGSQL" .OR. cEngine == "PGSQL64" .OR. cEngine == "POSTGRESQL"
      cComando := "ROLLBACK;"
   CASE cEngine == "DUCKDB" .OR. cEngine == "DUCKLAKE"
      cComando := "ROLLBACK;"
   ENDCASE
   RETURN cComando


// +--------------------------------------------------------------------
// +
// +    Funções RDD de Transação Refatoradas
// +
// +--------------------------------------------------------------------

STATIC FUNCTION ADO_TRANSBEGIN( nWA )
   LOCAL oConn := hb_RDDADOXGetConnection( nWA )
   IF oConn:State != adStateClosed
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         oConn:Execute( ADO_DIALETO_BEGIN( t_cEngine ) )
      RECOVER
         // Fallback de segurança: se o driver ODBC rejeitar SQL direto, usa o método OLE
         oConn:BeginTrans()
      END SEQUENCE
   ENDIF
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_TRANSCOMMIT( nWA )
   LOCAL oConn := hb_RDDADOXGetConnection( nWA )
   IF oConn:State != adStateClosed
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         oConn:Execute( ADO_DIALETO_COMMIT( t_cEngine ) )
      RECOVER
         oConn:CommitTrans()
      END SEQUENCE
   ENDIF
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_TRANSROLLBACK( nWA )
   LOCAL oConn := hb_RDDADOXGetConnection( nWA )
   IF oConn:State != adStateClosed
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         oConn:Execute( ADO_DIALETO_ROLLBACK( t_cEngine ) )
      RECOVER
         oConn:RollbackTrans()
      END SEQUENCE
   ENDIF
   RETURN HB_SUCCESS


// +--------------------------------------------------------------------
// +
// +    Static Function ADO_OPEN()
// +
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_OPEN( nWA, aOpenInfo )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL cName, aField, oError, nResult
   LOCAL oRecordSet, nTotalFields, n
   LOCAL cDataBase, cEXTENSAO, cDIRETORIO, cConnString

   cDataBase  := aOpenInfo[ UR_OI_NAME ]
   cNAME      := ""
   cEXTENSAO  := ""
   cDIRETORIO := ""

   IF Empty( aOpenInfo[ UR_OI_ALIAS ] )
      hb_FNameSplit( cDataBase, @cDIRETORIO, @cName, @CEXTENSAO )
      aOpenInfo[ UR_OI_ALIAS ] := cName
   ENDIF

   IF Empty( aOpenInfo[ UR_OI_CONNECT ] )
      aWAData[ WA_CONNECTION ] := win_oleCreateObject( "ADODB.Connection" )
      aWAData[ WA_TABLENAME ]  := t_cTableName
      aWAData[ WA_QUERY ]      := t_cQuery
      aWAData[ WA_USERNAME ]   := t_cUserName
      aWAData[ WA_PASSWORD ]   := t_cPassword
      aWAData[ WA_SERVER ]     := t_cServer
      aWAData[ WA_ENGINE ]     := t_cEngine
      aWAData[ WA_CONNOPEN ]   := .T.

      // Aciona a Fábrica de Conexões internalizada
      cConnString := ADO_BUILD_CONNECTION_STRING( cDataBase )
      
      nResult := HB_SUCCESS
      BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
         aWAData[ WA_CONNECTION ]:Open( cConnString )
      RECOVER USING oError
         UR_SUPER_ERROR( nWA, oError )
         nResult := HB_FAILURE
      END SEQUENCE

      IF nResult == HB_FAILURE
         RETURN HB_FAILURE
      ENDIF

   ELSE
      aWAData[ WA_CONNECTION ]         := win_oleAuto()
      aWAData[ WA_CONNECTION ] :__hObj := aOpenInfo[ UR_OI_CONNECT ] 
      aWAData[ WA_TABLENAME ]          := t_cTableName
      aWAData[ WA_QUERY ]              := t_cQuery
      aWAData[ WA_USERNAME ]           := t_cUserName
      aWAData[ WA_PASSWORD ]           := t_cPassword
      aWAData[ WA_SERVER ]             := t_cServer
      aWAData[ WA_ENGINE ]             := t_cEngine
      aWAData[ WA_CONNOPEN ]           := .F.
   ENDIF

   t_cQuery := ""

   IF Empty( aWAData[ WA_QUERY ] )
      aWAData[ WA_QUERY ] := "SELECT * FROM "
   ENDIF

   oRecordSet := win_oleCreateObject( "ADODB.Recordset" )

   IF oRecordSet == NIL
      oError             := ErrorNew()
      oError:GenCode     := EG_OPEN
      oError:SubCode     := 1001
      oError:Description := hb_langErrMsg( EG_OPEN )
      oError:FileName    := cDataBase
      oError:CanDefault  := .T.
      UR_SUPER_ERROR( nWA, oError )
      RETURN HB_FAILURE
   ENDIF

   oRecordSet:CursorType     := adOpenDynamic
   oRecordSet:CursorLocation := adUseClient
   oRecordSet:LockType       := adLockPessimistic
   
   IF aWAData[ WA_QUERY ] == "SELECT * FROM "
      oRecordSet:Open( aWAData[ WA_QUERY ] + aWAData[ WA_TABLENAME ], aWAData[ WA_CONNECTION ] )
   ELSE
      oRecordSet:Open( aWAData[ WA_QUERY ], aWAData[ WA_CONNECTION ] )
   ENDIF

   // ITEM 3: Carregamento nativo de Índices sem ADOX (em memória)
   aWAData[ WA_CATALOG ] := ADO_LOAD_INDEXES( aWAData[ WA_CONNECTION ], aWAData[ WA_TABLENAME ] )

   aWAData[ WA_RECORDSET ] := oRecordSet
   aWAData[ WA_BOF ]       := aWAData[ WA_EOF ] := .F.

   UR_SUPER_SETFIELDEXTENT( nWA, nTotalFields := oRecordSet:Fields:Count )

   FOR n := 1 TO nTotalFields
      aField                  := Array( UR_FI_SIZE )
      aField[ UR_FI_NAME ]    := oRecordSet:Fields( n - 1 ) :Name
      aField[ UR_FI_TYPE ]    := ADO_GETFIELDTYPE( oRecordSet:Fields( n - 1 ) :Type )
      aField[ UR_FI_TYPEEXT ] := 0
      // Agora passamos o objeto Field inteiro para calcular precisão e decimais!
      aField[ UR_FI_LEN ]     := ADO_GETFIELDSIZE( oRecordSet:Fields( n - 1 ), aField[ UR_FI_TYPE ] )
      aField[ UR_FI_DEC ]     := ADO_GETFIELDDEC( oRecordSet:Fields( n - 1 ), aField[ UR_FI_TYPE ] )
      UR_SUPER_ADDFIELD( nWA, aField )
   NEXT

   nResult := UR_SUPER_OPEN( nWA, aOpenInfo )

   IF nResult == HB_SUCCESS
      ADO_GOTOP( nWA )
   ENDIF

   RETURN nResult

// +--------------------------------------------------------------------
// +    Static Function ADO_BUILD_CONNECTION_STRING()
// +    Fábrica centralizada de strings ODBC/OLEDB com suporte a 32/64 bits
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_BUILD_CONNECTION_STRING( cDatabase )
   LOCAL cConn    := ""
   LOCAL cEngine  := t_cEngine
   LOCAL cServer  := t_cServer
   LOCAL cUser    := t_cUserName
   LOCAL cPass    := t_cPassword
   LOCAL cDir     := hb_FNameDir( cDatabase )
   LOCAL cSQLUSER := ""

   IF Right( AllTrim( cDir ), 1 ) != "\" .AND. Right( AllTrim( cDir ), 1 ) != "/"
      cDir += "\"
   ENDIF

   DO CASE
   // MDB/ACCESS (32 bits usa o Jet)
   CASE cEngine == "MDB" .OR. cEngine == "ACCESS"
      cConn := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDatabase + ";Mode=Share Deny None"
   
   // ACCDB/64 bits (usa o ACE OLEDB)
   CASE cEngine == "ACCDB" .OR. cEngine == "ACCDB64" .OR. cEngine == "MDB64" .OR. cEngine == "ACCESS64" .OR. cEngine == "ACEOLEDB"
      cConn := "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + cDatabase + ";Mode=Share Deny None"
   
   // SQLITE (Driver único)
   CASE cEngine == "SQLITE"
      cConn := "Driver={SQLite3 ODBC Driver};Database=" + cDatabase + ";"
      
   // DUCKDB (Driver único)
   CASE cEngine == "DUCKDB"
      cConn := "Driver={DuckDB Driver};Database=" + cDatabase + ";"
      
   // MYSQL / MARIADB (Separação 32 vs 64 bits)
   CASE cEngine == "MYSQL" 
      cConn := "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" + cServer + ";Database=" + cDatabase + ";Uid=" + cUser + ";Pwd=" + cPass + ";"
   CASE cEngine == "MYSQL64"
      cConn := "Driver={MySQL ODBC 9.0 ANSI Driver};Server=" + cServer + ";Database=" + cDatabase + ";Uid=" + cUser + ";Pwd=" + cPass + ";"
   CASE cEngine == "MARIADB"
      cConn := "Driver={MariaDB ODBC 3.2 Driver};Server=" + cServer + ";Database=" + cDatabase + ";Uid=" + cUser + ";Pwd=" + cPass + ";"
               
   // POSTGRESQL (Separação 32 vs 64 bits)
   CASE cEngine == "PGSQL" .OR. cEngine == "POSTGRESQL"
      cConn := "DRIVER={PostgreSQL ANSI};Server=" + cServer + ";Database=" + cDatabase + ";Uid=" + cUser + ";Pwd=" + cPass + ";ConnSettings=SET client_encoding TO 'WIN1252';"
   CASE cEngine == "PGSQL64"
      cConn := "DRIVER={PostgreSQL ANSI(x64)};Server=" + cServer + ";Database=" + cDatabase + ";Uid=" + cUser + ";Pwd=" + cPass + ";ConnSettings=SET client_encoding TO 'WIN1252';"
               
   // SQL SERVER / MSSQL
   CASE cEngine == "MSSQL" .OR. cEngine == "SQLSERVER" .OR. cEngine == "SQL"
      cSQLUSER := iif( Empty( cUser ), ";Trusted_Connection=True;", ";Uid=" + cUser + ";Pwd=" + cPass + ";" )
      cConn := "Provider=SQLOLEDB;Server=" + cServer + ";Database=" + cDatabase + cSQLUSER
      
   // FIREBIRD
   CASE cEngine == "FIREBIRD" .OR. cEngine == "FDB" .OR. cEngine == "GDB" .OR. cEngine == "IB"
      // Se preferir, pode-se injetar a chamada nativa DriverFirebird() se ela for movida para cá futuramente
      cConn := "DRIVER={Firebird ODBC Driver};UID=" + cUser + ";PWD=" + cPass + ";DBNAME=" + cDatabase
      
   // ORACLE
   CASE cEngine == "ORACLE" .OR. cEngine == "OCI"
      cConn := "Provider=MSDAORA.1;Persist Security Info=False;Data source=" + cDatabase + ";User ID=" + cUser + ";Password=" + cPass
      
   // PADRÕES DE ARQUIVO
   CASE cEngine == "PARADOX"
      cConn := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDir + ";Extended Properties=Paradox 5.x;"
      
   CASE cEngine == "DBASE"
      cConn := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDir + ";Extended Properties=dBASE IV;"
      
   CASE cEngine == "XLS"
      cConn := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDatabase + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'"
   ENDCASE

   RETURN cConn
   
   
   // +--------------------------------------------------------------------
// +    Static Function ADO_LOAD_INDEXES()
// +    Lê índices nativamente (adSchemaIndexes = 12) retornando um Array
// +--------------------------------------------------------------------
STATIC FUNCTION ADO_LOAD_INDEXES( oConn, cTableName )
   LOCAL oRs, aIndexes := {}
   LOCAL cIndexName

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      oRs := oConn:OpenSchema( 12 ) // 12 = adSchemaIndexes
      IF oRs != NIL .AND. oRs:State == 1
         DO WHILE !oRs:EOF
            // Garante que só pegará índices da tabela aberta
            IF Upper( AllTrim( oRs:Fields( "TABLE_NAME" ):Value ) ) == Upper( AllTrim( cTableName ) )
               cIndexName := oRs:Fields( "INDEX_NAME" ):Value
               IF !Empty( cIndexName ) .AND. AScan( aIndexes, {|x| Upper( x ) == Upper( cIndexName ) } ) == 0
                  AAdd( aIndexes, cIndexName )
               ENDIF
            ENDIF
            oRs:MoveNext()
         ENDDO
         oRs:Close()
      ENDIF
   RECOVER
   END SEQUENCE
   RETURN aIndexes
   
   
   // +--------------------------------------------------------------------
// +    Static Functions: GETROW, GETROWBLANK e PUTROW
// +    Manipulação direta de registros via Array tipado
// +--------------------------------------------------------------------

STATIC FUNCTION ADO_GETROW( nWA )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nTotalFld  := oRecordSet:Fields:Count
   LOCAL aRow       := Array( nTotalFld )
   LOCAL i, xVal

   IF aWAData[ WA_EOF ] .OR. oRecordSet:EOF .OR. oRecordSet:BOF
      RETURN ADO_GETROWBLANK( nWA )
   ENDIF

   FOR i := 1 TO nTotalFld
      xVal := oRecordSet:Fields( i - 1 ):Value
      IF ValType( xVal ) == "U"
         DO CASE
         CASE ADO_GETFIELDTYPE( oRecordSet:Fields( i - 1 ):Type ) == HB_FT_STRING
            xVal := Space( oRecordSet:Fields( i - 1 ):DefinedSize )
         CASE ADO_GETFIELDTYPE( oRecordSet:Fields( i - 1 ):Type ) == HB_FT_DATE
            xVal := hb_SToD()
         CASE ADO_GETFIELDTYPE( oRecordSet:Fields( i - 1 ):Type ) == HB_FT_TIMESTAMP
            xVal := hb_SToD()
         CASE ADO_GETFIELDTYPE( oRecordSet:Fields( i - 1 ):Type ) == HB_FT_LOGICAL
            xVal := .F.
         OTHERWISE
            xVal := 0
         ENDCASE
      ENDIF
      aRow[ i ] := xVal
   NEXT i

   RETURN aRow

STATIC FUNCTION ADO_GETROWBLANK( nWA )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nTotalFld  := oRecordSet:Fields:Count
   LOCAL aRow       := Array( nTotalFld )
   LOCAL i, nType

   FOR i := 1 TO nTotalFld
      nType := ADO_GETFIELDTYPE( oRecordSet:Fields( i - 1 ):Type )
      DO CASE
      CASE nType == HB_FT_STRING
         aRow[ i ] := Space( oRecordSet:Fields( i - 1 ):DefinedSize )
      CASE nType == HB_FT_LOGICAL
         aRow[ i ] := .F.
      CASE nType == HB_FT_DATE .OR. nType == HB_FT_TIMESTAMP
         aRow[ i ] := hb_SToD()
      CASE nType == HB_FT_INTEGER .OR. nType == HB_FT_LONG .OR. nType == HB_FT_DOUBLE
         aRow[ i ] := 0
      CASE nType == HB_FT_MEMO .OR. nType == HB_FT_OLE
         aRow[ i ] := ""
      OTHERWISE
         aRow[ i ] := ""
      ENDCASE
   NEXT i

   RETURN aRow

STATIC FUNCTION ADO_PUTROW( nWA, aRow )
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nTotalFld  := oRecordSet:Fields:Count
   LOCAL i

   IF aWAData[ WA_EOF ] .OR. oRecordSet:EOF
      RETURN HB_FAILURE
   ENDIF

   IF ValType( aRow ) != "A" .OR. Len( aRow ) < nTotalFld
      RETURN HB_FAILURE
   ENDIF

   BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
      FOR i := 1 TO nTotalFld
         IF !( oRecordSet:Fields( i - 1 ):Value == aRow[ i ] )
            oRecordSet:Fields( i - 1 ):Value := aRow[ i ]
         ENDIF
      NEXT i
   RECOVER
      RETURN HB_FAILURE
   END SEQUENCE

   RETURN HB_SUCCESS