/*
 * Firebird Open Source OLE DB driver
 *
 * SCHEMAROWSET.CPP - Source file for schema rowset related stuff.
 *
 * Distributable under LGPL license.
 * You may obtain a copy of the License at http://www.gnu.org/copyleft/lgpl.html
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * LGPL License for more details.
 *
 * This file was created by Ralph Curry.
 * All individual contributions remain the Copyright (C) of those
 * individuals. Contributors to this file are either listed here or
 * can be obtained from a CVS history command.
 *
 * All rights reserved.
 */

#include <windows.h>
#include <ibase.h>
#include "../inc/common.h"
#include "../inc/schemarowset.h"
#include "../inc/rowset.h"
#include "../inc/util.h"
#include "../inc/convert.h"

int sqlnamecmp(
		char	  *pszName1,
		char	  *pszName2,
		int		   nLen)
{
	if (pszName2[nLen] == ' ')
		pszName2[nLen] = '\0';
	return strcmp(pszName1, pszName2);
}

CSchemaRowsetHelper::CSchemaRowsetHelper( void )
{
	m_piscDbHandle = NULL;
	m_iscTransHandle = NULL;
	m_iscStmtHandle = NULL;
	m_pOutSqlda = NULL;
	m_pInSqlda = NULL;
	m_cMaxColumns = 0;
}

CSchemaRowsetHelper::~CSchemaRowsetHelper( void )
{
	if( m_iscTransHandle != NULL )
	{
		if( m_iscStatus[ 1 ] == 0 )
			isc_commit_transaction( m_iscStatus, &m_iscTransHandle );
		else
			isc_rollback_transaction( m_iscStatus, &m_iscTransHandle );
	}
	if( m_pInSqlda != NULL )
		delete[] m_pInSqlda;
	if( m_pOutSqlda != NULL )
		delete[] m_pOutSqlda;

	if( m_iscStmtHandle != NULL )
		isc_dsql_free_statement(m_iscStatus, &m_iscStmtHandle, DSQL_drop);
}
bool CSchemaRowsetHelper::Init( isc_db_handle *piscDbHandle )
{
	m_piscDbHandle = piscDbHandle;
	if( isc_start_transaction( m_iscStatus, &m_iscTransHandle, 1, m_piscDbHandle, 0, NULL ) )
	{
		return false;
	}
	return true;
}

void CSchemaRowsetHelper::SetCommand( const char *szCommand )
{
	m_cPredicates = 0;
	m_fNeedsWhere = true;
	m_strCommand = szCommand;
}

const char *CSchemaRowsetHelper::GetCommand( void )
{
	return m_strCommand.c_str();
}

void CSchemaRowsetHelper::SetColumnCount( short cMaxColumns )
{
	m_cMaxColumns = cMaxColumns;
}

bool CSchemaRowsetHelper::Append( const char *szText, int iLen )
{
	m_strCommand += " ";
	m_strCommand.append( szText, iLen );
	return true;
}

bool CSchemaRowsetHelper::Append( const char *szText )
{
	m_strCommand += " ";
	m_strCommand += szText;
	return true;
}

bool CSchemaRowsetHelper::AppendWhere( const char *szPredicate )
{
	if( szPredicate != NULL )
	{

		m_strCommand += m_fNeedsWhere ? " where " : " and ";
		m_strCommand += szPredicate;
	}

	m_fNeedsWhere = false;

	return true;
}

bool CSchemaRowsetHelper::AppendWhereParam( const char *szPredicate )
{
	m_cPredicates++;

	return AppendWhere( szPredicate );
}

bool CSchemaRowsetHelper::BindCol( short sColumn, short sqltype, short *psStatus, void *pData, short sLength )
{
	if( sqltype == IB_DEFAULT )
	{
		if( m_pOutSqlda == NULL ||
			m_pOutSqlda->sqln <= sColumn )
		{
			return false;
		}
		sqltype = m_pOutSqlda->sqlvar[ sColumn ].sqltype;
	}
	return Bind( BindTypeColumn, sColumn, sqltype, psStatus, pData, sLength );
}

bool CSchemaRowsetHelper::BindParam( short sColumn, short sqltype, short *psStatus, void *pData, short sLength )
{
	if( sColumn == IB_DEFAULT )
	{
		sColumn = m_cPredicates - 1;
	}
	return Bind( BindTypeParameter, sColumn, sqltype, psStatus, pData, sLength );
}

bool CSchemaRowsetHelper::Bind( BindTypeEnum eBindType, short sColumn, short sqltype, short *psStatus, void *pData, short sLength )
{
	XSQLVAR *pSqlvar = NULL;
	XSQLDA **ppSqlda = ( eBindType == BindTypeColumn ? &m_pOutSqlda : &m_pInSqlda );

	if( *ppSqlda == NULL ||
		sColumn > ( *ppSqlda )->sqln )
	{
		XSQLDA *pSqlda;

		pSqlda = (XSQLDA *)new BYTE[ XSQLDA_LENGTH( sColumn + 5 ) ];
		pSqlda->sqln = sColumn + 5;
		pSqlda->version = SQLDA_VERSION1;

		if( *ppSqlda != NULL )
		{
			memcpy( pSqlda->sqlvar, ( *ppSqlda )->sqlvar, ( *ppSqlda )->sqld );
			delete[] (BYTE *)(*ppSqlda);
		}

		*ppSqlda = pSqlda;
	}
	( *ppSqlda )->sqld = MAX( sColumn + 1, ( *ppSqlda )->sqld );
	pSqlvar = &( *ppSqlda )->sqlvar[ sColumn ];
	pSqlvar->sqldata = (char *)pData;
	pSqlvar->sqlind = psStatus;
	pSqlvar->sqllen = sLength;
	pSqlvar->sqlscale = 0;
	pSqlvar->sqlsubtype = 0;
	pSqlvar->sqltype = sqltype;

	return true;
}

bool CSchemaRowsetHelper::Exec( void )
{
	if( m_pOutSqlda == NULL )
	{
		if( m_cMaxColumns == 0 )
		{
			return false;
		}
		m_pOutSqlda = (XSQLDA *)new BYTE[ XSQLDA_LENGTH( m_cMaxColumns ) ];
		m_pOutSqlda->sqln    = m_cMaxColumns;
		m_pOutSqlda->version = SQLDA_VERSION1;
	}

	if( ( m_rc = isc_dsql_allocate_statement( m_iscStatus, m_piscDbHandle, &m_iscStmtHandle ) ) ||
		( m_rc = isc_dsql_prepare( m_iscStatus, &m_iscTransHandle, &m_iscStmtHandle, m_strCommand.length(), (char *)m_strCommand.c_str(), g_dso_usDialect, m_pOutSqlda ) ) ||
		( m_rc = isc_dsql_execute( m_iscStatus, &m_iscTransHandle, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda ) ) )
	{
		isc_rollback_transaction( m_iscStatus, &m_iscTransHandle );
		m_iscTransHandle = NULL;
		return false;
	}
	return true;
}

bool CSchemaRowsetHelper::Fetch( void )
{
	m_rc = isc_dsql_fetch( m_iscStatus, &m_iscStmtHandle, 1, m_pOutSqlda );
	return( m_rc == 0 );
}

HRESULT SchemaColumnsRowset(
	isc_db_handle  *piscDbHandle, 
	CComPolyObject<CInterBaseRowset> *pRowset,
	ULONG          cRestrictions, 
	const VARIANT  rgRestrictions[] )
{
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_COLUMNNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_COLUMNGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_COLUMNPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG,        IBCOLUMN_ORDINALPOSITION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_COLUMNHASDEFAULT )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,               128, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_COLUMNDEFAULT )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG,        IBCOLUMN_COLUMNFLAGS )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_ISNULLABLE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI2,  SQL_SHORT,       IBCOLUMN_DATATYPE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_TYPEGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_CHARACTERMAXIMUMLENGTH )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_CHARACTEROCTETLENGTH )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI2,  SQL_SHORT + 1,   IBCOLUMN_NUMERICPRECISION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I2,   SQL_SHORT + 1,   IBCOLUMN_NUMERICSCALE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_DATETIMEPRECISION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_CHARACTERSETCATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_CHARACTERSETSCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_CHARACTERSETNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_COLLATIONCATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_COLLATIONSCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_COLLATIONNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DOMAINCATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DOMAINSCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DOMAINNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                80, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DESCRIPTION )
	END_PREFETCHEDCOLUMNINFO
	PIBROW  pIbRow;
	ULONG   ulRestriction;
	ULONG   ulCharLength;
	ULONG   ulByteLength;
	short   sValue;
	long    lValue;
	short   rgsStatus[ 11 ];
	char   *pszTableName  = (char *)alloca( MAX_SQL_NAME_SIZE );
	char   *pszColumnName = (char *)alloca( MAX_SQL_NAME_SIZE );
	DWORD   dwFlags;
	short   sPosition;
	short   sNullFlag1;
	short   sNullFlag2;
	short   sUpdatable;
	WORD    wType;
	short   sSubType;
	short   sLength;
	short   sScale;
	short   sDimensions;
	USHORT  usPrecision;
	bool    fSuccess;

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );
	SchemaRowsetHelper.SetCommand( "select rf.RDB$RELATION_NAME, rf.RDB$FIELD_NAME, rf.RDB$FIELD_POSITION, rf.RDB$UPDATE_FLAG, rf.RDB$NULL_FLAG, f.RDB$NULL_FLAG, f.RDB$FIELD_TYPE, f.RDB$FIELD_SUB_TYPE, f.RDB$FIELD_LENGTH, f.RDB$FIELD_SCALE, f.RDB$DIMENSIONS  from RDB$RELATION_FIELDS rf join RDB$FIELDS f on rf.RDB$FIELD_SOURCE = f.RDB$FIELD_NAME" );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			return E_INVALIDARG;
		}

		switch( ulRestriction )
		{
		case 2:
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszTableName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "rf.RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
			}
			break;
		case 3:
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszColumnName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "rf.RDB$FIELD_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszColumnName, strlen( pszColumnName ) );
			}
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}
	SchemaRowsetHelper.Append( "order by 1, 3" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	SchemaRowsetHelper.SetColumnCount( 11 );
	SchemaRowsetHelper.Exec();
	SchemaRowsetHelper.BindCol( 10, IB_DEFAULT, &rgsStatus[10], &sDimensions,  sizeof( short ) );
	SchemaRowsetHelper.BindCol(  9, IB_DEFAULT, &rgsStatus[ 9], &sScale,       sizeof( short ) );
	SchemaRowsetHelper.BindCol(  8, IB_DEFAULT, &rgsStatus[ 8], &sLength,      sizeof( short ) );
	SchemaRowsetHelper.BindCol(  7, IB_DEFAULT, &rgsStatus[ 7], &sSubType,     sizeof( short ) );
	SchemaRowsetHelper.BindCol(  6, IB_DEFAULT, &rgsStatus[ 6], &wType,        sizeof( short ) );
	SchemaRowsetHelper.BindCol(  5, IB_DEFAULT, &rgsStatus[ 5], &sNullFlag2,   sizeof( short ) );
	SchemaRowsetHelper.BindCol(  4, IB_DEFAULT, &rgsStatus[ 4], &sNullFlag1,   sizeof( short ) );
	SchemaRowsetHelper.BindCol(  3, IB_DEFAULT, &rgsStatus[ 3], &sUpdatable,   sizeof( short ) );
	SchemaRowsetHelper.BindCol(  2, IB_DEFAULT, &rgsStatus[ 2], &sPosition,    sizeof( short ) );
	SchemaRowsetHelper.BindCol(  1, IB_DEFAULT, &rgsStatus[ 1], pszColumnName, MAX_SQL_NAME_SIZE );
	SchemaRowsetHelper.BindCol(  0, IB_DEFAULT, &rgsStatus[ 0], pszTableName,  MAX_SQL_NAME_SIZE );

	do
	{
		pIbRow = pRowset->m_contained.AllocRow( );
		
		if( pIbRow != NULL &&
			( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
		{
			if( rgsStatus[ 4 ] == 0 || rgsStatus[ 5 ] == 0 )
				rgsStatus[ 4 ] = 0;
			ulByteLength = ulCharLength = sLength;
			BlrTypeToOledbType( (short)wType, sSubType, &wType, &usPrecision, &sScale, &dwFlags, &ulCharLength, &ulByteLength );
			if( rgsStatus[ 10 ] == 0 )
			{
				wType = DBTYPE_ARRAY;
			}

			if( sUpdatable == 1 )
				dwFlags |= DBCOLUMNFLAGS_WRITEUNKNOWN;
			if( rgsStatus[4] == -1 )
				dwFlags |= ( DBCOLUMNFLAGS_ISNULLABLE | DBCOLUMNFLAGS_MAYBENULL );

			pRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, SQL_TEXT,   MAX_SQL_NAME_SIZE,  pszTableName  );
			pRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow, SQL_TEXT,   MAX_SQL_NAME_SIZE,  pszColumnName );
			pRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT,       0, (void *)&(lValue = sPosition + 1) );
			pRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT,       0, (void *)&(sValue = 0) );
			pRowset->m_contained.SetPrefetchedColumnData( 8, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 9, pIbRow, IB_DEFAULT,       0, (void *)&(dwFlags) );
			pRowset->m_contained.SetPrefetchedColumnData(10, pIbRow, IB_DEFAULT,       0, (void *)&(sValue = ( rgsStatus[4] == -1 ) ) );
			pRowset->m_contained.SetPrefetchedColumnData(11, pIbRow, IB_DEFAULT,       0, (void *)&wType );
			pRowset->m_contained.SetPrefetchedColumnData(12, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(13, pIbRow, IB_DEFAULT,       0, (void *)&ulCharLength );
			pRowset->m_contained.SetPrefetchedColumnData(14, pIbRow, IB_DEFAULT,       0, (void *)(wType == DBTYPE_STR || wType == DBTYPE_BYTES ? &ulByteLength : NULL ) );
			pRowset->m_contained.SetPrefetchedColumnData(15, pIbRow, IB_DEFAULT,       0, (void *)(wType == DBTYPE_NUMERIC ? &usPrecision : NULL ) );
			pRowset->m_contained.SetPrefetchedColumnData(16, pIbRow, IB_DEFAULT,       0, (void *)(wType == DBTYPE_NUMERIC ? &sScale : NULL ) );
			pRowset->m_contained.SetPrefetchedColumnData(17, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(18, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(19, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(20, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(21, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(22, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(23, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(24, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(25, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(26, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(27, pIbRow, IB_DEFAULT, IB_NULL, NULL );

			pRowset->m_contained.InsertPrefetchedRow( pIbRow );
		}
	} while( fSuccess );
	pRowset->m_contained.RestartPosition( NULL );

	if( pIbRow == NULL )
		return E_OUTOFMEMORY;
	
	return S_OK;
}

HRESULT SchemaTablesRowset(
	isc_db_handle  *piscDbHandle, 
	CComPolyObject<CInterBaseRowset> *pRowset,
	ULONG          cRestrictions, 
	const VARIANT  rgRestrictions[] )
{
	HRESULT      hr = S_OK;
	ULONG        ulRestriction;
	PIBROW       pIbRow = NULL;
	char         szType[2];
	short        sSystemFlag;
	short        rgsStatus[3];
	short        sTableType;
	bool         fSuccess;
	char        *pszTableName = (char *)alloca( MAX_SQL_NAME_SIZE );
	VARIANT     *pvtTableName = NULL;
	VARIANT     *pvtTableType = NULL;
	const char  *rgszTableTypes[] = { "VIEW", "TABLE", "SYSTEM VIEW", "SYSTEM TABLE" };
	const char   szViewSelect[]  = "select rdb$relation_name, rdb$system_flag, 'V' from rdb$relations where RDB$VIEW_SOURCE is not NULL";
	const char   szTableSelect[] = "select rdb$relation_name, rdb$system_flag, 'T' from rdb$relations where RDB$VIEW_SOURCE is NULL";
/*
select rdb$relation_name, rdb$system_flag, rdb$view_source 
from   rdb$relations
*/
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,    MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,    MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,    MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                   16, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TABLETYPE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                   38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_TABLEGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                   80, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DESCRIPTION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                    0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_TABLEPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,  sizeof(DBTIMESTAMP), DBTYPE_DATE, SQL_DATE + 1,    IBCOLUMN_DATECREATED )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,  sizeof(DBTIMESTAMP), DBTYPE_DATE, SQL_DATE + 1,    IBCOLUMN_DATEMODIFIED )
	END_PREFETCHEDCOLUMNINFO

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			return E_INVALIDARG;
		}

		switch( ulRestriction )
		{
		case 2:
			{
				ULONG cbLength;
				pvtTableName = (VARIANT *)&rgRestrictions[ ulRestriction ];
				GetStringFromVariant( (char **)&pszTableName, MAX_SQL_NAME_SIZE, &cbLength, pvtTableName );
			}
			break;
		case 3:
			pvtTableType = (VARIANT *)&rgRestrictions[ ulRestriction ];
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}

	if( pvtTableName == NULL &&
		pvtTableType == NULL )
	{
		SchemaRowsetHelper.SetCommand( szTableSelect );
		SchemaRowsetHelper.Append( "union all" );
		SchemaRowsetHelper.Append( szViewSelect );
	}
	else
	{
		if( pvtTableType != NULL )
		{
			bool fSystem = false;

			switch( HashString( pvtTableType->bstrVal ) )
			{
				case 0x000069d5: //TABLE
					SchemaRowsetHelper.SetCommand( szTableSelect );
					break;
				case 0x00001b7b: //VIEW
					SchemaRowsetHelper.SetCommand( szViewSelect );
					break;
				case 0x1c06f9d5: //SYSTEM TABLE
					fSystem = true;
					SchemaRowsetHelper.SetCommand( szTableSelect );
					break;
				case 0x0701bf7b: //SYSTEM VIEW
					fSystem = true;
					SchemaRowsetHelper.SetCommand( szViewSelect );
					break;
			}
			SchemaRowsetHelper.AppendWhere( NULL );
			if( fSystem )
			{
				SchemaRowsetHelper.AppendWhere( "RDB$SYSTEM_FLAG = 1" );
			}
			else
			{
				SchemaRowsetHelper.AppendWhere( "RDB$SYSTEM_FLAG = 0" );
			}

			if( pvtTableName != NULL )
			{
				SchemaRowsetHelper.AppendWhereParam( "RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
			}
		}
		else
		{
			SchemaRowsetHelper.SetCommand( szTableSelect );
			SchemaRowsetHelper.AppendWhere( NULL );
			SchemaRowsetHelper.AppendWhereParam( "RDB$RELATION_NAME=?" );
			SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
			SchemaRowsetHelper.Append( "union all" );
			SchemaRowsetHelper.Append( szViewSelect );
			SchemaRowsetHelper.AppendWhereParam( "RDB$RELATION_NAME=?" );
			SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
		}
	}

	SchemaRowsetHelper.Append( "order by 2 desc, 3, 1" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	SchemaRowsetHelper.SetColumnCount( 3 );
	SchemaRowsetHelper.Exec();
	SchemaRowsetHelper.BindCol( 2, IB_DEFAULT, &rgsStatus[2], szType, sizeof( szType ) );
	SchemaRowsetHelper.BindCol( 1, IB_DEFAULT, &rgsStatus[1], &sSystemFlag, sizeof( sSystemFlag ) );
	SchemaRowsetHelper.BindCol( 0, IB_DEFAULT, &rgsStatus[0], pszTableName, MAX_SQL_NAME_SIZE );
	do
	{
		pIbRow = pRowset->m_contained.AllocRow( );
		
		if( pIbRow != NULL &&
			( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
		{
			sTableType = ( szType[ 0 ] == 'T' ) + ( ( sSystemFlag != 0 ) * 2 );
			_ASSERTE( sTableType >= 0 && sTableType <= 3 );
			pRowset->m_contained.SetPrefetchedColumnData(0, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(1, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(2, pIbRow, SQL_TEXT,   MAX_SQL_NAME_SIZE, pszTableName );
			pRowset->m_contained.SetPrefetchedColumnData(3, pIbRow, SQL_TEXT,              IB_NTS, (void *)rgszTableTypes[ sTableType ]);
			pRowset->m_contained.SetPrefetchedColumnData(4, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(5, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(6, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(7, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(8, pIbRow, IB_DEFAULT,           IB_NULL, NULL );

			pRowset->m_contained.InsertPrefetchedRow( pIbRow );
		}
	} while( fSuccess );
	pRowset->m_contained.RestartPosition( NULL );

	if( pIbRow == NULL )
		return E_OUTOFMEMORY;

	return S_OK;
}

HRESULT SchemaProviderTypesRowset(
	isc_db_handle  *piscDbHandle, 
	CComPolyObject<CInterBaseRowset> *pRowset,
	ULONG          cRestrictions, 
	const VARIANT  rgRestrictions[] )
{
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TYPENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI2,  SQL_SHORT,       IBCOLUMN_DATATYPE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG,        IBCOLUMN_COLUMNSIZE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 4, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_LITERALPREFIX )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 4, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_LITERALSUFFIX )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                16, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_CREATEPARAMS )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_ISNULLABLE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_CASESENSITIVE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG,        IBCOLUMN_SEARCHABLE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_VARYING + 1, IBCOLUMN_UNSIGNEDATTRIBUTE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_FIXEDPRECSCALE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_AUTOUNIQUEVALUE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_LOCALTYPENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_I2,   SQL_VARYING + 1, IBCOLUMN_MINIMUMSCALE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_I2,   SQL_VARYING + 1, IBCOLUMN_MAXIMUMSCALE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_GUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TYPELIB )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_VERSION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_ISLONG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_BESTMATCH )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 5, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_ISFIXEDLENGTH )
	END_PREFETCHEDCOLUMNINFO

	const struct PROVIDERTYPES {
		BYTE   bVersion;
		char  *szTypeName;
		USHORT sType;
		ULONG  ulColumnSize;
		char  *szLiteralPrefix;
		char  *szLiteralSuffix;
		char  *szCreateParams;
		short  sCaseSensitive;
		ULONG uiSearchable;
		char  *szUnsigned;
		short  sFixedPrecScale;
		char  *szMinScale;
		char  *szMaxScale;
		short  sIsLong;
		short  sBestMatch;
		short  sIsFixedLength;
	} rgProviderTypes[] =
	{
		{ 3, "smallint",           DBTYPE_I2,             5, NULL, NULL, NULL,               0, DB_ALL_EXCEPT_LIKE, "True" , 0, NULL, NULL, 0, 1, 1 },
		{ 3, "integer",            DBTYPE_I4,            10, NULL, NULL, NULL,               0, DB_ALL_EXCEPT_LIKE, "True" , 0, NULL, NULL, 0, 1, 1 },
		{ 3, "float",              DBTYPE_R4,             7, NULL, NULL, NULL,               0, DB_ALL_EXCEPT_LIKE, "False", 0, NULL, NULL, 0, 1, 1 },
		{ 3, "double precision",   DBTYPE_R8,            15, NULL, NULL, NULL,               0, DB_ALL_EXCEPT_LIKE, "False", 0, NULL, NULL, 0, 1, 1 },
		{ 3, "char",               DBTYPE_STR,        32767, "'",  "'",  "length",           1, DB_SEARCHABLE,      NULL,    0, NULL, NULL, 0, 0, 1 },
		{ 3, "varchar",            DBTYPE_STR,        32767, "'",  "'",  "max length",       1, DB_SEARCHABLE,      NULL,    0, NULL, NULL, 0, 1, 0 },
		{ 3, "nchar",              DBTYPE_STR,        16535, "'",  "'",  "length",           1, DB_SEARCHABLE,      NULL,    0, NULL, NULL, 0, 0, 1 },
		{ 3, "nchar varying",      DBTYPE_STR,        16535, "'",  "'",  "max length",       1, DB_SEARCHABLE,      NULL,    0, NULL, NULL, 0, 0, 0 },
		{ 3, "blob sub_type text", DBTYPE_STR,   2147483647, NULL, NULL, NULL,               1, DB_UNSEARCHABLE,    NULL,    0, NULL, NULL, 1, 0, 0 },
		{ 3, "blob",               DBTYPE_BYTES, 2147483647, NULL, NULL, NULL,               1, DB_UNSEARCHABLE,    NULL,    0, NULL, NULL, 1, 1, 0 },
		{ 2, "date",               DBTYPE_DBDATE,        10, "'",  "'",  NULL,               0, DB_ALL_EXCEPT_LIKE, NULL,    0, NULL, NULL, 0, 1, 1 },
		{ 2, "time",               DBTYPE_DBTIME,         8, "'",  "'",  NULL,               0, DB_ALL_EXCEPT_LIKE, NULL,    0, NULL, NULL, 0, 1, 1 },
		{ 2, "timestamp",          DBTYPE_DBTIMESTAMP,   23, "'",  "'",  NULL,               0, DB_ALL_EXCEPT_LIKE, NULL,    0, NULL, NULL, 0, 1, 1 },
		{ 2, "decimal",            DBTYPE_NUMERIC,       19, NULL, NULL, "precision, scale", 0, DB_ALL_EXCEPT_LIKE, "False", 1, "0",  "19", 0, 0, 1 },
		{ 2, "numeric",            DBTYPE_NUMERIC,       19, NULL, NULL, "precision, scale", 0, DB_ALL_EXCEPT_LIKE, "False", 1, "0",  "19", 0, 1, 1 },
		{ 1, "date",               DBTYPE_DBTIMESTAMP,   23, "'",  "'",  NULL,               0, DB_ALL_EXCEPT_LIKE, NULL,    0, NULL, NULL, 0, 1, 1 },
		{ 1, "decimal",            DBTYPE_NUMERIC,        9, NULL, NULL, "precision, scale", 0, DB_ALL_EXCEPT_LIKE, "False", 0, "0",   "9", 0, 0, 1 },
		{ 1, "numeric",            DBTYPE_NUMERIC,        9, NULL, NULL, "precision, scale", 0, DB_ALL_EXCEPT_LIKE, "False", 0, "0",   "9", 0, 1, 1 },
	};
	VARIANT rgvtRestrictions[ 2 ];
	VARTYPE usType;
	PIBROW  pIbRow = NULL;
	ULONG   ulRestriction;
	int     i;
	short   sBool;
	bool    fNoResults = false;
	BYTE    bVersion = (g_dso_usDialect == 3) + 1;

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	VariantInit( &rgvtRestrictions[ 0 ] );
	VariantInit( &rgvtRestrictions[ 1 ] );
	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_NULL:
			fNoResults = true;
		}

		switch( ulRestriction )
		{
		case 0:
			usType = VT_I2;
			break;
		case 1:
			usType = VT_BOOL;
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
		if( FAILED( VariantChangeTypeEx( &rgvtRestrictions[ ulRestriction ], (VARIANT *)&rgRestrictions[ ulRestriction ], NULL, 0, usType ) ) )
		{
			return E_INVALIDARG;
		}
	}

	if( fNoResults )
	{
		return S_OK;
	}

	for (i = 0;i < sizeof( rgProviderTypes ) / sizeof( struct PROVIDERTYPES ); i++ )
	{
		if( rgProviderTypes[ i ].bVersion & bVersion )
		{
			if( ( rgvtRestrictions[ 0 ].vt != VT_EMPTY &&
				  rgvtRestrictions[ 0 ].iVal != rgProviderTypes[ i ].sType ) ||
				( rgvtRestrictions[ 1 ].vt != VT_EMPTY &&
				  ( rgvtRestrictions[ 1 ].boolVal == -1 ) != rgProviderTypes[ i ].sBestMatch ) )  
			{
				continue;
			}
			pIbRow = pRowset->m_contained.AllocRow( );
			pRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szTypeName );
			pRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sType );
			pRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].ulColumnSize );
			pRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szLiteralPrefix );
			pRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szLiteralSuffix );
			pRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szCreateParams );
			pRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT,       0, (void *)&( sBool = 1 ) );
			pRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sCaseSensitive );
			pRowset->m_contained.SetPrefetchedColumnData( 8, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].uiSearchable );
			pRowset->m_contained.SetPrefetchedColumnData( 9, pIbRow, SQL_TEXT,   IB_NULL, rgProviderTypes[ i ].szUnsigned );
			pRowset->m_contained.SetPrefetchedColumnData(10, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sFixedPrecScale );
			pRowset->m_contained.SetPrefetchedColumnData(11, pIbRow, IB_DEFAULT,       0, (void *)&( sBool = 0 ) );
			pRowset->m_contained.SetPrefetchedColumnData(12, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szMinScale );
			pRowset->m_contained.SetPrefetchedColumnData(13, pIbRow, SQL_TEXT,    IB_NTS, rgProviderTypes[ i ].szMaxScale );
			pRowset->m_contained.SetPrefetchedColumnData(14, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(15, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(16, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(17, pIbRow, IB_DEFAULT, IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(18, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sIsLong );
			pRowset->m_contained.SetPrefetchedColumnData(19, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sBestMatch );
			pRowset->m_contained.SetPrefetchedColumnData(20, pIbRow, IB_DEFAULT,       0, (void *)&rgProviderTypes[ i ].sIsFixedLength );

			pRowset->m_contained.InsertPrefetchedRow( pIbRow );
		}
	}
	pRowset->m_contained.RestartPosition( NULL );

	return S_OK;
}

HRESULT GetKeyColumnInfo(
		XSQLDA			  *pDescSqlda,
		IBCOLUMNMAP		  *pColumnMap,
		isc_db_handle	  *piscDbHandle,
		isc_tr_handle	  *piscTransHandle)
{
static char SqlSelect[]  = "select i.RDB$SEGMENT_COUNT, s.RDB$FIELD_NAME, i.RDB$RELATION_NAME "
                           "from RDB$INDEX_SEGMENTS s inner join (RDB$INDICES i inner join RDB$RELATION_CONSTRAINTS c on i.RDB$INDEX_NAME = c.RDB$INDEX_NAME) on i.RDB$INDEX_NAME = s.RDB$INDEX_NAME "
                           "where RDB$CONSTRAINT_TYPE = 'PRIMARY KEY' and i.RDB$INDEX_INACTIVE is NULL and i.RDB$RELATION_NAME in ";

	HRESULT		   hr = S_OK;
	DWORD		   dwFlags = 0;
	int			   i;
	int			   nFound;
	int			   nUniqueTableNames = 0;
	XSQLVAR		 **rgpUniqueTableNames = NULL;	//array of pointers to sqlvars with unique table names
	XSQLVAR		  *pSqlVar = pDescSqlda->sqlvar;
	XSQLVAR		  *pSqlVarEnd = pSqlVar + pDescSqlda->sqld;
	IBCOLUMNMAP	  *pColumn;
	IBCOLUMNMAP	  *rgpKeyColumns[16];	//array of key columns, assumes we won't have an index with over 16 columns

	if (pDescSqlda->sqld < 1)
		return S_OK;

	//Allocate enough memory to track the unique table names.  We know sqld is an upper bound.
	rgpUniqueTableNames = (XSQLVAR **)new char [ sizeof(XSQLVAR *) * pDescSqlda->sqld ];

	//We know the first table name is unique
	rgpUniqueTableNames[ nUniqueTableNames++ ] = pSqlVar++;

	//Look for any other unique table names
	for( ; pSqlVar < pSqlVarEnd; pSqlVar++ )
	{
		for( i = 0; i < nUniqueTableNames; i++ )
		{
			if( !strcmp(pSqlVar->relname, rgpUniqueTableNames[ i ]->relname ) )
			{
				break;
			}
		}
		if( i == nUniqueTableNames )
		{
			//got one, add it to the array
			rgpUniqueTableNames[nUniqueTableNames++] = pSqlVar;
		}
	}

	short rgsStatus[3];
	short sKeyCount;
	char  szColumnName[ MAX_SQL_NAME_SIZE ];
	char  szTableName[ MAX_SQL_NAME_SIZE ];
	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );

	SchemaRowsetHelper.SetCommand( SqlSelect );

	//add all the table names in single quotes and any closing characters
	SchemaRowsetHelper.Append( "(" );

	for (i = 0; i < nUniqueTableNames; i++)
	{
		if (i > 0)
		{
			SchemaRowsetHelper.Append( ",?" );
		}
		else
		{
			SchemaRowsetHelper.Append( "?" );
		}
		SchemaRowsetHelper.AppendWhereParam( NULL );
		SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, &sKeyCount, rgpUniqueTableNames[i]->relname, rgpUniqueTableNames[i]->relname_length );
	}

	SchemaRowsetHelper.Append( ")" );

	if( rgpUniqueTableNames != NULL )
	{
		delete[] rgpUniqueTableNames;
	}

	sKeyCount = 0;
	SchemaRowsetHelper.SetColumnCount( 3 );
	if( ! SchemaRowsetHelper.Exec() )
	{
		hr = E_FAIL;
		goto cleanup;
	}

	SchemaRowsetHelper.BindCol( 0, IB_DEFAULT, &rgsStatus[ 0 ], (void *)&sKeyCount,   sizeof( sKeyCount ) );
	SchemaRowsetHelper.BindCol( 1, IB_DEFAULT, &rgsStatus[ 1 ], (void *)szColumnName, sizeof( szColumnName ) );
	SchemaRowsetHelper.BindCol( 2, IB_DEFAULT, &rgsStatus[ 2 ], (void *)szTableName,  sizeof( szTableName ) );

	i = 0;	//number of fields left in the current index
	pSqlVarEnd = pDescSqlda->sqlvar + pDescSqlda->sqld;
	
	//get all the index field information, and see if we have those columns
	//in the select list
	while( SchemaRowsetHelper.Fetch() )
	{
		if (i == 0)
		{
			//we're starting a new index, so initialize these
			nFound = 0;
			i = sKeyCount;
		}

		//loop through all the fields in the select list
		for (pSqlVar = pDescSqlda->sqlvar; pSqlVar < pSqlVarEnd; pSqlVar++)
		{
			//check if the column name matches and also the table name
			//if nUniqueTableNames != 1
			if (!sqlnamecmp(pSqlVar->sqlname, szColumnName, pSqlVar->sqlname_length) &&
				(nUniqueTableNames == 1 ||
				 !sqlnamecmp(pSqlVar->relname, szTableName, pSqlVar->relname_length)))
			{
				//get the right IbColumnMap
				pColumn = pColumnMap + (pSqlVar - pDescSqlda->sqlvar);

				//add it to the array of columns that comprise the key
				rgpKeyColumns[ nFound++ ] = pColumn;

				//do we have all the fields we need for this index?
				if( sKeyCount == nFound )
				{
					//mark the key column bit for all the key columns
					while( nFound-- )
					{
						rgpKeyColumns[ nFound ]->dwFlags |= DBCOLUMNFLAGS_KEYCOLUMN;
					}

					break;
				}
			}
		}
		i--;
	}

cleanup:
	return hr;
}

#ifdef OPTIONAL_SCHEMA_ROWSETS
HRESULT SchemaForeignKeysRowset(
	isc_db_handle  *piscDbHandle,
	CComPolyObject<CInterBaseRowset> *pRowset, 
	ULONG           cRestrictions, 
	const VARIANT   rgRestrictions[] )
{
	HRESULT      hr = S_OK;
	ULONG        ulRestriction;
	PIBROW       pIbRow = NULL;
	short        rgsStatus[ 11 ];
	bool         fSuccess;
	short        sDeferrability;
	long         lOrdinal;
	char         szPKName[ MAX_SQL_NAME_SIZE ];
	char         szFKName[ MAX_SQL_NAME_SIZE ];
	char         szPKColumnName[ MAX_SQL_NAME_SIZE ];
	char         szFKColumnName[ MAX_SQL_NAME_SIZE ];
	char         szUpdateRule[ 16 ];
	char         szDeleteRule[ 16 ];
	char         szDeferrable[ 4 ];
	char         szDeferred[ 4 ];
	char        *pszPKTableName = (char *)alloca( MAX_SQL_NAME_SIZE );
	char        *pszFKTableName = (char *)alloca( MAX_SQL_NAME_SIZE );
/*
select rcpk.RDB$RELATION_NAME, ispk.RDB$FIELD_NAME, rcfk.RDB$RELATION_NAME, (select RDB$FIELD_NAME from RDB$INDEX_SEGMENTS isfk where isfk.RDB$INDEX_NAME = rcfk.RDB$INDEX_NAME and isfk.RDB$FIELD_POSITION = ispk.RDB$FIELD_POSITION), ispk.RDB$FIELD_POSITION, rc.RDB$UPDATE_RULE, rc.RDB$DELETE_RULE, rcpk.RDB$CONSTRAINT_NAME, rcfk.RDB$CONSTRAINT_NAME, rcfk.RDB$DEFERRABLE, rcfk.RDB$INITIALLY_DEFERRED
from RDB$INDEX_SEGMENTS ispk join 
RDB$RELATION_CONSTRAINTS rcpk on ispk.RDB$INDEX_NAME = rcpk.RDB$INDEX_NAME join
RDB$REF_CONSTRAINTS rc on rcpk.RDB$CONSTRAINT_NAME = rc.RDB$CONST_NAME_UQ join
RDB$RELATION_CONSTRAINTS rcfk on rc.RDB$CONSTRAINT_NAME = rcfk.RDB$CONSTRAINT_NAME
order by 3,5
*/
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PKTABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PKTABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_PKTABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_PKCOLUMNNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_PKCOLUMNGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_PKCOLUMNPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_FKTABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_FKTABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_FKTABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_FKCOLUMNNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_FKCOLUMNGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_FKCOLUMNPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_ORDINAL )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                16, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_UPDATERULE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                16, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_DELETERULE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_PKNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_FKNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I2,   SQL_SHORT + 1,   IBCOLUMN_DEFERRABILITY )
	END_PREFETCHEDCOLUMNINFO

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );
	SchemaRowsetHelper.SetCommand( "select rcpk.RDB$RELATION_NAME, ispk.RDB$FIELD_NAME, rcfk.RDB$RELATION_NAME, (select RDB$FIELD_NAME from RDB$INDEX_SEGMENTS isfk where isfk.RDB$INDEX_NAME = rcfk.RDB$INDEX_NAME and isfk.RDB$FIELD_POSITION = ispk.RDB$FIELD_POSITION), ispk.RDB$FIELD_POSITION, rc.RDB$UPDATE_RULE, rc.RDB$DELETE_RULE, rcpk.RDB$CONSTRAINT_NAME, rcfk.RDB$CONSTRAINT_NAME, rcfk.RDB$DEFERRABLE, rcfk.RDB$INITIALLY_DEFERRED from RDB$INDEX_SEGMENTS ispk join RDB$RELATION_CONSTRAINTS rcpk on ispk.RDB$INDEX_NAME = rcpk.RDB$INDEX_NAME join RDB$REF_CONSTRAINTS rc on rcpk.RDB$CONSTRAINT_NAME = rc.RDB$CONST_NAME_UQ join RDB$RELATION_CONSTRAINTS rcfk on rc.RDB$CONSTRAINT_NAME = rcfk.RDB$CONSTRAINT_NAME" );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			return E_INVALIDARG;
		}

		switch( ulRestriction )
		{
		case 2:
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszPKTableName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "rcpk.RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszPKTableName, strlen( pszPKTableName ) );
			}
			break;
		case 5:
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszFKTableName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "rcfk.RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszFKTableName, strlen( pszFKTableName ) );
			}
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}

	SchemaRowsetHelper.Append( "order by 3, 5" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	SchemaRowsetHelper.SetColumnCount( 11 );
	SchemaRowsetHelper.Exec();
	SchemaRowsetHelper.BindCol( 10, IB_DEFAULT, &rgsStatus[10], szDeferred,     sizeof( szDeferred ) );
	SchemaRowsetHelper.BindCol(  9, IB_DEFAULT, &rgsStatus[ 9], szDeferrable,   sizeof( szDeferrable ) );
	SchemaRowsetHelper.BindCol(  8, IB_DEFAULT, &rgsStatus[ 8], szFKName,       sizeof( szFKName ) );
	SchemaRowsetHelper.BindCol(  7, IB_DEFAULT, &rgsStatus[ 7], szPKName,       sizeof( szPKName ) );
	SchemaRowsetHelper.BindCol(  6, IB_DEFAULT, &rgsStatus[ 6], szDeleteRule,   sizeof( szDeleteRule ) );
	SchemaRowsetHelper.BindCol(  5, IB_DEFAULT, &rgsStatus[ 5], szUpdateRule,   sizeof( szUpdateRule ) );
	SchemaRowsetHelper.BindCol(  4, SQL_LONG,   &rgsStatus[ 4], &lOrdinal,      sizeof( lOrdinal ) );
	SchemaRowsetHelper.BindCol(  3, IB_DEFAULT, &rgsStatus[ 3], szFKColumnName, sizeof( szFKColumnName ) );
	SchemaRowsetHelper.BindCol(  2, IB_DEFAULT, &rgsStatus[ 2], pszFKTableName,  MAX_SQL_NAME_SIZE );
	SchemaRowsetHelper.BindCol(  1, IB_DEFAULT, &rgsStatus[ 1], szPKColumnName, sizeof( szPKColumnName ) );
	SchemaRowsetHelper.BindCol(  0, IB_DEFAULT, &rgsStatus[ 0], pszPKTableName,  MAX_SQL_NAME_SIZE );
	do
	{
		pIbRow = pRowset->m_contained.AllocRow( );
		
		if( pIbRow != NULL &&
			( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
		{
			pRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, SQL_TEXT,        MAX_SQL_NAME_SIZE, pszPKTableName );
			pRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow, SQL_TEXT, sizeof( szPKColumnName ), szPKColumnName );
			pRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData( 8, pIbRow, SQL_TEXT,        MAX_SQL_NAME_SIZE, pszFKTableName );
			pRowset->m_contained.SetPrefetchedColumnData( 9, pIbRow, SQL_TEXT, sizeof( szFKColumnName ), szFKColumnName );
			pRowset->m_contained.SetPrefetchedColumnData(10, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(11, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(12, pIbRow, IB_DEFAULT,                      0, (void *)&(++lOrdinal) );
			pRowset->m_contained.SetPrefetchedColumnData(13, pIbRow, SQL_TEXT,                  IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(14, pIbRow, SQL_TEXT,                  IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(15, pIbRow, SQL_TEXT,       sizeof( szPKName ), szPKName );
			pRowset->m_contained.SetPrefetchedColumnData(16, pIbRow, SQL_TEXT,       sizeof( szFKName ), szFKName );
			
			if( szDeferrable[ 0 ] == 'N' )
			{
				sDeferrability = DBPROPVAL_DF_NOT_DEFERRABLE;
			}
			else if( szDeferred[ 0 ] == 'N' )
			{
				sDeferrability = DBPROPVAL_DF_INITIALLY_IMMEDIATE;
			}
			else
			{
				sDeferrability = DBPROPVAL_DF_INITIALLY_DEFERRED;
			}

			pRowset->m_contained.SetPrefetchedColumnData(17, pIbRow, IB_DEFAULT,                      0, (void *)&(sDeferrability) );

			pRowset->m_contained.InsertPrefetchedRow( pIbRow );
		}
	} while( fSuccess );
	pRowset->m_contained.RestartPosition( NULL );

	if( pIbRow == NULL )
		return E_OUTOFMEMORY;

	return S_OK;
}

HRESULT SchemaIndexesRowset(
	isc_db_handle  *piscDbHandle,
	CComPolyObject<CInterBaseRowset> *pRowset, 
	ULONG           cRestrictions, 
	const VARIANT   rgRestrictions[] )
{
	HRESULT      hr = S_OK;
	ULONG        ulRestriction;
	PIBROW       pIbRow = NULL;
	short        rgsStatus[ 6 ];
	bool         fSuccess;
	short        sValue;
	short        sUnique;
	short        sPrimaryKey;
	short        sOrdinal;
	long         lValue;
	char         szColumnName[ MAX_SQL_NAME_SIZE ];
	char        *pszIndexName = (char *)alloca( MAX_SQL_NAME_SIZE );
	char        *pszTableName = (char *)alloca( MAX_SQL_NAME_SIZE );
	bool         fMatchNone = false;
/*
select
    RDB$RELATION_NAME
  , ri.RDB$INDEX_NAME
  , (select cast (1 as SMALLINT ) from RDB$RELATION_CONSTRAINTS rrc where rrc.RDB$CONSTRAINT_TYPE='PRIMARY KEY' and RDB$INDEX_NAME=ri.RDB$INDEX_NAME)
  , RDB$UNIQUE_FLAG
  , RDB$FIELD_POSITION
  , RDB$FIELD_NAME
from RDB$INDICES ri join
rdb$INDEX_SEGMENTS ris on ri.RDB$INDEX_NAME = ris.RDB$INDEX_NAME
where (ri.RDB$INDEX_INACTIVE is NULL or RDB$INDEX_INACTIVE = 0)
order by 4, 2, 5
*/
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_INDEXCATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_INDEXSCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_INDEXNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT + 1,   IBCOLUMN_PRIMARYKEY )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_UNIQUE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_CLUSTERED )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI2,  SQL_SHORT,       IBCOLUMN_TYPE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I4,   SQL_LONG,        IBCOLUMN_FILLFACTOR )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_INITIALSIZE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_NULLS )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_SORTBOOKMARKS )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_BOOL, SQL_SHORT,       IBCOLUMN_AUTOUPDATE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I4,   SQL_LONG,        IBCOLUMN_NULLCOLLATION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG,        IBCOLUMN_ORDINALPOSITION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_COLUMNNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_COLUMNGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_COLUMNPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I2,   SQL_SHORT + 1,   IBCOLUMN_COLLATION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI8,  SQL_LONG + 1,    IBCOLUMN_CARDINALITY )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I4,   SQL_LONG + 1,    IBCOLUMN_PAGES )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_FILTERCONDITION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I2,   SQL_SHORT + 1,   IBCOLUMN_INTEGRATED )
	END_PREFETCHEDCOLUMNINFO

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );
	SchemaRowsetHelper.SetCommand( 
		"select "
		"    RDB$RELATION_NAME "
		"  , ri.RDB$INDEX_NAME "
		"  , (select cast (1 as SMALLINT) from RDB$RELATION_CONSTRAINTS rrc where rrc.RDB$CONSTRAINT_TYPE='PRIMARY KEY' and RDB$INDEX_NAME=ri.RDB$INDEX_NAME) "
		"  , RDB$UNIQUE_FLAG "
		"  , RDB$FIELD_POSITION "
		"  , RDB$FIELD_NAME "
		"from RDB$INDICES ri join "
		"rdb$INDEX_SEGMENTS ris on ri.RDB$INDEX_NAME = ris.RDB$INDEX_NAME "
		"where (ri.RDB$INDEX_INACTIVE is NULL or RDB$INDEX_INACTIVE = 0)" );

	SchemaRowsetHelper.AppendWhere( NULL );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			if( ulRestriction != 3 )
			{
				return E_INVALIDARG;
			}
		}

		switch( ulRestriction )
		{
		//case 0: //TABLE_CATALOG
		//case 1: //TABLE_SCHEMA
		case 2: //INDEX_NAME
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszIndexName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "ri.RDB$INDEX_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszIndexName, strlen( pszIndexName ) );
			}
			break;
		case 3: //TYPE
			{
				VARIANT tVariant[1];
				VariantInit( tVariant );
				if( VariantChangeType( tVariant, (VARIANT *)&rgRestrictions[ ulRestriction ], 0, VT_I2 ) != S_OK )
				{
					return E_INVALIDARG;
				}
				if( V_I2( tVariant ) != DBPROPVAL_IT_BTREE )
				{
					fMatchNone = true;
				}
				VariantClear( tVariant );
			}
			break;
		case 4: //TABLE_NAME
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszTableName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "ri.RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
			}
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}

	SchemaRowsetHelper.Append( "order by 2,4,5" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	if( ! fMatchNone )
	{
		SchemaRowsetHelper.SetColumnCount( 6 );
		SchemaRowsetHelper.Exec();
		SchemaRowsetHelper.BindCol(  5, IB_DEFAULT, &rgsStatus[ 5], szColumnName,  sizeof( szColumnName ) );
		SchemaRowsetHelper.BindCol(  4, IB_DEFAULT, &rgsStatus[ 4], &sOrdinal,     sizeof( sOrdinal ) );
		SchemaRowsetHelper.BindCol(  3, IB_DEFAULT, &rgsStatus[ 3], &sUnique,      sizeof( sUnique ) );
		SchemaRowsetHelper.BindCol(  2, IB_DEFAULT, &rgsStatus[ 2], &sPrimaryKey,  sizeof( sPrimaryKey ) );
		SchemaRowsetHelper.BindCol(  1, IB_DEFAULT, &rgsStatus[ 1], pszIndexName,  MAX_SQL_NAME_SIZE );
		SchemaRowsetHelper.BindCol(  0, IB_DEFAULT, &rgsStatus[ 0], pszTableName,  MAX_SQL_NAME_SIZE );
		do
		{
			pIbRow = pRowset->m_contained.AllocRow( );
			
			if( pIbRow != NULL &&
				( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
			{
				pRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, SQL_TEXT,        MAX_SQL_NAME_SIZE, pszTableName );
				pRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, SQL_TEXT,        MAX_SQL_NAME_SIZE, pszIndexName );
				pRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT,                      0, &sPrimaryKey );
				pRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT,                      0, &sUnique );
				pRowset->m_contained.SetPrefetchedColumnData( 8, pIbRow, IB_DEFAULT,                      0, &( sValue = 0 ) );
				pRowset->m_contained.SetPrefetchedColumnData( 9, pIbRow, IB_DEFAULT,                      0, &( sValue = DBPROPVAL_IT_BTREE ) );
				pRowset->m_contained.SetPrefetchedColumnData(10, pIbRow, IB_DEFAULT,                      0, &( lValue = 0 ) );
				pRowset->m_contained.SetPrefetchedColumnData(11, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(12, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(13, pIbRow, IB_DEFAULT,                      0, &( sValue = 0 ) );
				pRowset->m_contained.SetPrefetchedColumnData(14, pIbRow, IB_DEFAULT,                      0, &( sValue = ~0 ) );
				pRowset->m_contained.SetPrefetchedColumnData(15, pIbRow, IB_DEFAULT,                      0, &( lValue = DBPROPVAL_NC_END ) );
				pRowset->m_contained.SetPrefetchedColumnData(16, pIbRow, SQL_LONG,                        0, &( lValue = sOrdinal ) );
				pRowset->m_contained.SetPrefetchedColumnData(17, pIbRow, SQL_TEXT,   sizeof( szColumnName ), szColumnName );
				pRowset->m_contained.SetPrefetchedColumnData(18, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(19, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(20, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(21, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(22, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(23, pIbRow, IB_DEFAULT,                IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData(24, pIbRow, IB_DEFAULT,                      0, &( sValue = ~0 ) );
				
				pRowset->m_contained.InsertPrefetchedRow( pIbRow );
			}
		} while( fSuccess );
		pRowset->m_contained.RestartPosition( NULL );

		if( pIbRow == NULL )
			return E_OUTOFMEMORY;
	}

	return S_OK;
}

HRESULT SchemaPrimaryKeysRowset(
	isc_db_handle  *piscDbHandle, 
	CComPolyObject<CInterBaseRowset> *pRowset,
	ULONG          cRestrictions, 
	const VARIANT  rgRestrictions[] )
{
	HRESULT      hr = S_OK;
	ULONG        ulRestriction;
	PIBROW       pIbRow = NULL;
	short        rgsStatus[4];
	bool         fSuccess;
	long         lOrdinal;
	char        *pszTableName = (char *)alloca( MAX_SQL_NAME_SIZE );
	char         szColumnName[ MAX_SQL_NAME_SIZE ];
	char         szConstraintName[ MAX_SQL_NAME_SIZE ];
/*
select rdb$relation_name, rdb$field_name, rdb$field_position, rdb$constraint_name
from   rdb$relation_constraints rc join 
       rdb$index_segments rif on rc.rdb$index_name = rif.rdb$index_name
where  rdb$constraint_type = 'PRIMARY KEY';
*/
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_TABLESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_TABLENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_COLUMNNAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                38, DBTYPE_GUID, SQL_TEXT + 1,    IBCOLUMN_COLUMNGUID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_COLUMNPROPID )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_UI4,  SQL_LONG + 1,    IBCOLUMN_ORDINAL )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PKNAME )
	END_PREFETCHEDCOLUMNINFO

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );
	SchemaRowsetHelper.SetCommand( "select rdb$relation_name, rdb$field_name, rdb$field_position, rdb$constraint_name from rdb$relation_constraints rc join rdb$index_segments rif on rc.rdb$index_name = rif.rdb$index_name and rdb$constraint_type = 'PRIMARY KEY'" );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			return E_INVALIDARG;
		}

		switch( ulRestriction )
		{
		case 2:
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszTableName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "RDB$RELATION_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszTableName, strlen( pszTableName ) );
			}
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}

	SchemaRowsetHelper.Append( "order by 1" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	SchemaRowsetHelper.SetColumnCount( 4 );
	SchemaRowsetHelper.Exec();
	SchemaRowsetHelper.BindCol( 3, IB_DEFAULT, &rgsStatus[3], szConstraintName, sizeof( szConstraintName ) );
	SchemaRowsetHelper.BindCol( 2, SQL_LONG,   &rgsStatus[2], &lOrdinal, sizeof( lOrdinal ) );
	SchemaRowsetHelper.BindCol( 1, IB_DEFAULT, &rgsStatus[1], szColumnName, sizeof( szColumnName ) );
	SchemaRowsetHelper.BindCol( 0, IB_DEFAULT, &rgsStatus[0], pszTableName, MAX_SQL_NAME_SIZE );
	do
	{
		pIbRow = pRowset->m_contained.AllocRow( );
		
		if( pIbRow != NULL &&
			( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
		{
			pRowset->m_contained.SetPrefetchedColumnData(0, pIbRow, IB_DEFAULT,                    IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(1, pIbRow, IB_DEFAULT,                    IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(2, pIbRow, SQL_TEXT,            MAX_SQL_NAME_SIZE, pszTableName );
			pRowset->m_contained.SetPrefetchedColumnData(3, pIbRow, SQL_TEXT,       sizeof( szColumnName ), szColumnName );
			pRowset->m_contained.SetPrefetchedColumnData(4, pIbRow, IB_DEFAULT,                    IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(5, pIbRow, IB_DEFAULT,                    IB_NULL, NULL );
			pRowset->m_contained.SetPrefetchedColumnData(6, pIbRow, IB_DEFAULT,                          0, (void *)&(++lOrdinal) );
			pRowset->m_contained.SetPrefetchedColumnData(7, pIbRow, SQL_TEXT,   sizeof( szConstraintName ), szConstraintName );

			pRowset->m_contained.InsertPrefetchedRow( pIbRow );
		}
	} while( fSuccess );
	pRowset->m_contained.RestartPosition( NULL );

	if( pIbRow == NULL )
		return E_OUTOFMEMORY;

	return S_OK;
}

HRESULT SchemaProceduresRowset(
	isc_db_handle  *piscDbHandle,
	CComPolyObject<CInterBaseRowset> *pRowset, 
	ULONG           cRestrictions, 
	const VARIANT   rgRestrictions[] )
{
	HRESULT      hr = S_OK;
	ULONG        ulRestriction;
	PIBROW       pIbRow = NULL;
	short        rgsStatus[ 1 ];
	bool         fSuccess;
	short        sValue;
	char        *pszProcedureName = (char *)alloca( MAX_SQL_NAME_SIZE );
	bool         fMatchNone = false;
/*
select RDB$PROCEDURE_NAME 
from RDB$PROCEDURES
order by 1
*/
	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PROCEDURECATALOG )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PROCEDURESCHEMA )
		PREFETCHEDCOLUMNINFO_ENTRY( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR, SQL_VARYING,     IBCOLUMN_PROCEDURENAME )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_I2,   SQL_SHORT,       IBCOLUMN_PROCEDURETYPE )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                80, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_PROCEDUREDEFINITION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                80, DBTYPE_WSTR, SQL_VARYING + 1, IBCOLUMN_DESCRIPTION )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_DATE, SQL_DATE + 1,    IBCOLUMN_DATECREATED )
		PREFETCHEDCOLUMNINFO_ENTRY( 0,                 0, DBTYPE_DATE, SQL_DATE + 1,    IBCOLUMN_DATEMODIFIED )
	END_PREFETCHEDCOLUMNINFO

	CSchemaRowsetHelper SchemaRowsetHelper;
	SchemaRowsetHelper.Init( piscDbHandle );
	SchemaRowsetHelper.SetCommand( 
		"select RDB$PROCEDURE_NAME "
		"from RDB$PROCEDURES " );

	for (ulRestriction = 0; ulRestriction < cRestrictions; ulRestriction++ )
	{
		switch( rgRestrictions[ ulRestriction ].vt )
		{
		case VT_EMPTY:
			continue;
		case VT_BSTR:
			break;
		default:
			if( ulRestriction != 3 )
			{
				return E_INVALIDARG;
			}
		}

		switch( ulRestriction )
		{
		//case 0: //PROCEDURE_CATALOG
		//case 1: //PROCEDURE_SCHEMA
		case 2: //PROCEDURE_NAME
			{
				ULONG cbLength;
				GetStringFromVariant( (char **)&pszProcedureName, MAX_SQL_NAME_SIZE, &cbLength, (VARIANT *)&rgRestrictions[ ulRestriction ] );
				SchemaRowsetHelper.AppendWhereParam( "RDB$PROCEDURE_NAME=?" );
				SchemaRowsetHelper.BindParam( IB_DEFAULT, SQL_TEXT, 0, pszProcedureName, strlen( pszProcedureName ) );
			}
			break;
		case 3: //PROCEDURE_TYPE
			{
				VARIANT tVariant[1];
				VariantInit( tVariant );
				if( VariantChangeType( tVariant, (VARIANT *)&rgRestrictions[ ulRestriction ], 0, VT_I2 ) != S_OK )
				{
					return E_INVALIDARG;
				}
				if( V_I2( tVariant ) != DB_PT_PROCEDURE )
				{
					fMatchNone = true;
				}
				VariantClear( tVariant );
			}
			break;
		default:
			return DB_E_NOTSUPPORTED;
		}
	}

	SchemaRowsetHelper.Append( "order by 1" );

	pRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ), NULL );

	if( ! fMatchNone )
	{
		SchemaRowsetHelper.SetColumnCount( 1 );
		SchemaRowsetHelper.Exec();
		SchemaRowsetHelper.BindCol(  0, IB_DEFAULT, &rgsStatus[ 0], pszProcedureName,  MAX_SQL_NAME_SIZE );
		do
		{
			pIbRow = pRowset->m_contained.AllocRow( );
			
			if( pIbRow != NULL &&
				( fSuccess = SchemaRowsetHelper.Fetch() ) == true )
			{
				pRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, SQL_TEXT,   MAX_SQL_NAME_SIZE, pszProcedureName );
				pRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow, IB_DEFAULT,                 0, &( sValue = DB_PT_PROCEDURE ) );
				pRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				pRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT,           IB_NULL, NULL );
				
				pRowset->m_contained.InsertPrefetchedRow( pIbRow );
			}
		} while( fSuccess );
		pRowset->m_contained.RestartPosition( NULL );

		if( pIbRow == NULL )
			return E_OUTOFMEMORY;
	}
	return S_OK;
}
#endif //OPTIONAL_SCHEMA_ROWSETS
