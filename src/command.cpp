/*
 * Firebird Open Source OLE DB driver
 *
 * COMMAND.CPP - Source file for OLE DB command object methods.
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

#include "../inc/common.h"
#include "../inc/resource.h"
#include "../inc/iboledb.h"
#include "../inc/property.h"
#include "../inc/schemarowset.h"
#include "../inc/rowset.h"
#include "../inc/command.h"
#include "../inc/session.h"
#include "../inc/dso.h"
#include "../inc/smartlock.h"
#include "../inc/convert.h"
#include "../inc/util.h"
#include "../inc/blob.h"

CInterBaseCommand::CInterBaseCommand() :
	m_pOwningSession(NULL),
	m_pInterBaseError(NULL),
	m_iscStmtHandle(NULL),
	m_piscTransHandle(NULL)
{
	DbgFnTrace("CInterBaseCommand::CInterBaseCommand()");

	InitializeCriticalSection(&m_cs);

	//Internal state variables
	m_eCommandState = CommandStateUninitialized;
	m_cOpenRowsets = 0;
	m_fExecuting = false;
	m_fHaveKeyColumns = false;

	//Internal pointers
	m_pszCommand = NULL;
	m_pInSqlda = NULL;
	m_pOutSqlda = NULL;
	m_pIbColumnMap = NULL;
	m_pIbParamMap = NULL;

	//Accessor related stuff
	m_cIbAccessorCount = 0;
	m_pIbAccessorsToRelease = NULL;

}

CInterBaseCommand::~CInterBaseCommand()
{
	IBACCESSOR	  *pIbAccessor;
	IBPROPERTY	  *pIbProperties;
	ULONG		   cIbProperties;

	DbgFnTrace("CInterBaseCommand::~CInterBaseCommand()");

	if ((pIbProperties = _getIbProperties(DBPROPSET_ROWSET, &cIbProperties)) != NULL)
		ClearIbProperties(pIbProperties, cIbProperties);

	if( m_piscTransHandle != NULL )
	{
		m_pOwningSession->ReleaseTransHandle( &m_piscTransHandle );
	}

	if (m_pOwningSession)
		m_pOwningSession->Release();

	while (m_pIbAccessorsToRelease)
	{
		pIbAccessor = m_pIbAccessorsToRelease;
		m_pIbAccessorsToRelease = pIbAccessor->pNext;
		delete[] pIbAccessor;
	}

	if( this->m_iscStmtHandle != NULL )
		isc_dsql_free_statement(m_iscStatus, &m_iscStmtHandle, DSQL_drop);

	if (m_pszCommand)
		delete[] m_pszCommand;

	if (m_pIbColumnMap)
		delete[] m_pIbColumnMap;

	if (m_pIbParamMap)
		delete[] m_pIbParamMap;

	if (m_pOutSqlda)
		delete[] m_pOutSqlda;

	if (m_pInSqlda)
		delete[] m_pInSqlda;

	if( m_pInterBaseError )
		delete m_pInterBaseError;

	DeleteCriticalSection(&m_cs);
}

HRESULT CInterBaseCommand::Initialize(
		CInterBaseSession	  *pOwningSession, 
		isc_db_handle		  *piscDbHandle, 
		bool				   fUserCommand )
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::Initialize()");
	
	m_pOwningSession = pOwningSession;
	m_piscDbHandle = piscDbHandle;
	m_fUserCommand = fUserCommand;
	m_pOwningSession->AddRef();
	m_eCommandState = CommandStateInitial;

	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
// CInterBaseCommand::DeleteRowset
///////////////////////////////////////////////////////////////////////////////
// Called by the rowset when it is released.
// Either close the corsor and give the statement handle back to the Command 
// object, or drop the statement handle if the Command object already has a
// statement handle.
///////////////////////////////////////////////////////////////////////////////
HRESULT CInterBaseCommand::DeleteRowset(
		isc_stmt_handle	   iscStmtHandle)
{
	CSmartLock	sm(&m_cs);
	
	DbgFnTrace("CInterBaseCommand::DeleteRowset()");

	if( iscStmtHandle != NULL )
	{
	 	if( m_iscStmtHandle == NULL)
		{
			isc_dsql_free_statement(m_iscStatus, &iscStmtHandle, DSQL_close);
			m_iscStmtHandle = iscStmtHandle;
		}
		else
		{
			isc_dsql_free_statement(m_iscStatus, &iscStmtHandle, DSQL_drop);
		}
	}
	m_cOpenRowsets--;

	return S_OK;
}

HRESULT CInterBaseCommand::BeginMethod(
	REFIID	   riid)
{
	if (!(m_pInterBaseError ||
		 (m_pInterBaseError = new CInterBaseError)))
		return E_OUTOFMEMORY;

	m_pInterBaseError->ClearError(riid);

	return S_OK;
}

HRESULT CInterBaseCommand::MakeColumnMap(
		XSQLDA		  *pSqlda,
		IBCOLUMNMAP	 **ppColumnMap,
		ULONG		  *pcbBuffer)
{
	int		   nSize;
	ULONG	   i;
	ULONG	   cCount = pSqlda->sqld;
	ULONG	   cbBuffer = 0;
	IBCOLUMNMAP	  *pColumnMap;

	DbgFnTrace("CInterBaseCommand::MakeColumnMap()");

	pColumnMap = new IBCOLUMNMAP[pSqlda->sqld];
	if (!pColumnMap)
		return E_OUTOFMEMORY;

	for (i = 0; i < cCount; i++)
	{
		pColumnMap[i].columnid = DB_NULLID;
		pColumnMap[i].dwFlags = 0;
		pColumnMap[i].dwSpecialFlags = 0;
		pColumnMap[i].pvExtra = NULL;
		pColumnMap[i].iOrdinal = i + 1;
		pColumnMap[i].bPrecision = 0;
		pColumnMap[i].bScale = 0;

		//Figure out the wType and length related dwFlags
		if (ClosestOleDbType(
				pSqlda->sqlvar[i].sqltype,
				pSqlda->sqlvar[i].sqlsubtype,
				pSqlda->sqlvar[i].sqllen,
				pSqlda->sqlvar[i].sqlscale,
				&(pColumnMap[i].wType),
				&(pColumnMap[i].ulColumnSize),
				&(pColumnMap[i].bPrecision),
				&(pColumnMap[i].bScale),
				&(pColumnMap[i].dwFlags)) != DBSTATUS_S_OK)
			continue;

		pColumnMap[i].obData = cbBuffer;
		nSize = pSqlda->sqlvar[i].sqllen;
		switch( BASETYPE( pSqlda->sqlvar[i].sqltype ) )
		{
		case SQL_VARYING:
			nSize += 2;
			break;
		case SQL_TEXT:
			nSize += 1;
		}

		cbBuffer += ROUNDUP( nSize, sizeof(void *) );
		pColumnMap[i].obStatus = cbBuffer;
		cbBuffer += sizeof(void *);
		nSize = MultiByteToWideChar(
				CP_ACP,
				MB_PRECOMPOSED ,
				pSqlda->sqlvar[i].aliasname,
				pSqlda->sqlvar[i].aliasname_length,
				pColumnMap[i].wszAlias,
				MAX_SQL_NAME_SIZE);
		_ASSERT(nSize < MAX_SQL_NAME_SIZE);
		pColumnMap[i].wszAlias[nSize] = L'\0';
		nSize = MultiByteToWideChar(
				CP_ACP,
				MB_PRECOMPOSED ,
				pSqlda->sqlvar[i].sqlname,
				pSqlda->sqlvar[i].sqlname_length,
				pColumnMap[i].wszName,
				MAX_SQL_NAME_SIZE);
		_ASSERT(nSize < MAX_SQL_NAME_SIZE);
		pColumnMap[i].wszName[ nSize ] = L'\0';
	}

	*ppColumnMap = pColumnMap;
	*pcbBuffer = cbBuffer;
	return S_OK;
}

HRESULT CInterBaseCommand::DescribeBindings(void)
{
	HRESULT		   hr;
	ISC_STATUS	   sc;
	XSQLVAR       *pSqlvar;
	XSQLVAR       *pSqlvarEnd;

	DbgFnTrace("CInterBaseCommand::DescribeBindings()");

	if (!m_pInSqlda)
	{
		if (!(m_pInSqlda = (XSQLDA *)new char[XSQLDA_LENGTH(DEFAULT_COLUMN_COUNT)]))
			return E_OUTOFMEMORY;

		m_pInSqlda->sqln = DEFAULT_COLUMN_COUNT;
		m_pInSqlda->version = SQLDA_VERSION1;
	}
	
	if (sc = isc_dsql_describe_bind(m_iscStatus, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda))
		return m_pInterBaseError->PostError(m_iscStatus);

	if (m_pInSqlda->sqld > m_pInSqlda->sqln)
	{
		int	   n = m_pInSqlda->sqld;
		
		delete[] m_pInSqlda;

		if (!(m_pInSqlda = (XSQLDA *)new char[XSQLDA_LENGTH(n)]))
			return E_OUTOFMEMORY;
		
		m_pInSqlda->sqln = n;
		m_pInSqlda->version = SQLDA_VERSION1;
		
		if (sc = isc_dsql_describe_bind(m_iscStatus, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda))
			return m_pInterBaseError->PostError(m_iscStatus);
	}

	pSqlvar = m_pInSqlda->sqlvar;
	pSqlvarEnd = pSqlvar + m_pInSqlda->sqld;

	for( ; pSqlvar < pSqlvarEnd; pSqlvar++ )
	{
		if( BASETYPE( pSqlvar->sqltype ) == SQL_VARYING )
			pSqlvar->sqllen += 2;
	}
	if (FAILED(hr = MakeColumnMap(m_pInSqlda, &m_pIbParamMap, &m_cbParamBuffer)))
		return hr;

	m_cParams = m_pInSqlda->sqld;
	
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
// CInterBaseCommand::GetIbRow
///////////////////////////////////////////////////////////////////////////////
// Returns a buffer big enough to hold the fetched data for the current
// prepared command.  These buffers are pooled in the Command object for use
// by the Rowset object.  This should give us some efficiency if a prepared
// command is executed multiple times.
///////////////////////////////////////////////////////////////////////////////
/*
HRESULT CInterBaseCommand::GetIbRow(PIBROW *ppIbRow)
{
	CSmartLock sm(&m_cs);
	BYTE *pTemp;

	DbgFnTrace("CInterBaseRowset::GethIbRow()");

	if (m_pFreeIbRows)
	{
		_ASSERTE( m_pFreeIbRows->cbBytes >= m_cbFreeIbRow );
		*ppIbRow = m_pFreeIbRows;
		m_pFreeIbRows = m_pFreeIbRows->pNext;
	}
	else
	{
		pTemp = new BYTE[sizeof(IBROW) + m_cbFreeIbRow];
		if (!pTemp)
			return E_OUTOFMEMORY;
		((PIBROW)pTemp)->pvData = pTemp + sizeof(IBROW);
		((PIBROW)pTemp)->cbBytes = m_cbFreeIbRow;
		((PIBROW)pTemp)->pFree = NULL;
		((PIBROW)pTemp)->pNext = NULL;
		*ppIbRow = (PIBROW)pTemp;
//		((PIBROW)pTemp)->pFree = m_pIbRowsToRelease;
//		m_pIbRowsToRelease = (PIBROW)pTemp;
	}

	SetSqldaPointers(m_pOutSqlda, m_pIbColumnMap, *ppIbRow);

	return S_OK;
}
*/

///////////////////////////////////////////////////////////////////////////////
// CInterBaseCommand::GetIbValue
///////////////////////////////////////////////////////////////////////////////
// Used by the Command object during execution of an EXECUTE PROCEDURE
// statement or by the Rowset object during a GetData call.
// This method converts the data value from its native InterBase form to
// the OLE DB type requested in the Accessor.
///////////////////////////////////////////////////////////////////////////////
DBSTATUS CInterBaseCommand::GetIbValue(
		DBBINDING	  *pDbBinding,
		BYTE		  *pbData,
		XSQLVAR		  *pSqlVar,
		BYTE		  *pSqlData,
		short		   SqlInd)
{
	short		   SqlType = pSqlVar->sqltype;
	BYTE		   bPrecision = 0;
	BYTE		   bScale = 0;
	DBSTATUS	   wStatus = DBSTATUS_S_OK;
	DWORD		   dwLen = pDbBinding->cbMaxLen;

	if( NULLABLE( SqlType ) && SqlInd )
	{
		wStatus = DBSTATUS_S_ISNULL;
		dwLen = 0;
		goto skip_value;
	}

	if( ! ( pDbBinding->dwPart & ( DBPART_VALUE | DBPART_LENGTH ) ) )
	{
		goto skip_value;
	}

	switch ( BASETYPE( SqlType ) )
	{
	case SQL_BLOB:
		{
			isc_blob_handle	   iscBlobHandle = NULL;
			CComObject<CInterBaseBlobStream> *pIbBlobStream = NULL;

			if( ( isc_open_blob2(
					m_iscStatus, 
					m_piscDbHandle, 
					m_piscTransHandle, 
					&iscBlobHandle, 
					(ISC_QUAD *)pSqlData, 
					0, 
					NULL ) == 0 ) &&
				( ( pIbBlobStream = new CComObject<CInterBaseBlobStream> ) != NULL ) )
			{
				pIbBlobStream->AddRef();
				pIbBlobStream->Initialize(iscBlobHandle, pSqlVar->sqlsubtype == isc_blob_text);

				if( pDbBinding->dwPart & DBPART_LENGTH &&
					! (pDbBinding->dwPart & DBPART_VALUE ) )
				{
					BYTE	  *pbBuffer;
					DWORD	   cbBytesRead;

					pbBuffer = (BYTE *)alloca(STACK_BUFFER_SIZE);
					dwLen = 0;
					do
					{
						pIbBlobStream->Read( pbBuffer, STACK_BUFFER_SIZE, &cbBytesRead );
						dwLen += cbBytesRead;
					} while ( STACK_BUFFER_SIZE == cbBytesRead );
				}
				else if( pDbBinding->wType == DBTYPE_IUNKNOWN )
				{
					_ASSERTE( pDbBinding->dwPart & DBPART_VALUE );
					if( FAILED( pIbBlobStream->QueryInterface( IID_ISequentialStream, (void **)(pbData + pDbBinding->obValue) ) ) )
					{
						wStatus = DBSTATUS_E_CANTCREATE;
					}
				}
				else
				{
					pIbBlobStream->Read( pbData + pDbBinding->obValue, dwLen, &dwLen );
					/*
					if( FAILED( g_pIDataConvert->DataConvert(
							DBTYPE_STR,
							pDbBinding->wType,
							dwLen,
							&dwLen,
							pbData + pDbBinding->obValue,
							pvData,
							pDbBinding->cbMaxLen,
							wStatus,
							&wStatus,
							NULL,
							NULL,
							0 ) ) )
					{
						wStatus = DBSTATUS_E_CANTCONVERTVALUE;
					}
					*/
				}
			}
			else
			{
				wStatus = DBSTATUS_E_CANTCREATE;
			}

			if( pIbBlobStream != NULL )
				pIbBlobStream->Release();

			if( wStatus == DBSTATUS_E_CANTCREATE )
			{
				if( iscBlobHandle != NULL )
					isc_close_blob( m_iscStatus, &iscBlobHandle );
			}
		}
		break;

	case SQL_ARRAY:
		wStatus = GetArray(
				m_piscDbHandle,
				m_piscTransHandle,
				pSqlVar,
				pSqlData,
				&SqlInd,
				pDbBinding->wType,
				pDbBinding->dwPart & DBPART_VALUE ? pbData + pDbBinding->obValue : NULL,
				&dwLen,
				&(pDbBinding->dwFlags),
				&(m_pIbColumnMap[pDbBinding->iOrdinal].pvExtra));
		break;
	
	default:
		InterBaseToOleDbType(
				SqlType,
				pSqlVar->sqlsubtype,
				pSqlVar->sqllen,
				pSqlVar->sqlscale,
				pSqlData,
				&SqlInd,
				pDbBinding->wType,
				&dwLen,
				&bPrecision,
				&bScale,
				pDbBinding->dwPart & DBPART_VALUE ? pbData + pDbBinding->obValue : NULL,
				&wStatus);
	}

skip_value:
	if( pDbBinding->dwPart & DBPART_STATUS )
	{
		*(DBSTATUS *)(pbData + pDbBinding->obStatus) = wStatus;
	}

	if( pDbBinding->dwPart & DBPART_LENGTH )
	{
		*(DWORD *)(pbData + pDbBinding->obLength) = dwLen;
	}

	return wStatus;
}

CRITICAL_SECTION *CInterBaseCommand::GetCriticalSection(void)
{
	return &m_cs;
}

/*
///////////////////////////////////////////////////////////////////////////////
// CInterBaseCommand::ReleaseIbRow
///////////////////////////////////////////////////////////////////////////////
// Give the fetch buffer back to the command for use by a later fetch.
///////////////////////////////////////////////////////////////////////////////
HRESULT CInterBaseCommand::ReleaseIbRow(
		IBROW	  *pIbRow)
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::ReleaseIbRow()");

	if( pIbRow->cbBytes >= this->m_cbFreeIbRow )
	{
		pIbRow->pNext = m_pFreeIbRows;
		m_pFreeIbRows = pIbRow;
	}
	else
	{
		delete (BYTE *)pIbRow;
	}

	return S_OK;
}
*/

HRESULT WINAPI CInterBaseCommand::InternalQueryInterface(
	CInterBaseCommand *			pThis,
	const _ATL_INTMAP_ENTRY*	pEntries, 
	REFIID						iid, 
	void**						ppvObject)
{
	HRESULT    hr;
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::InternalQueryInterface()");

	DbgPrintGuid( iid );

	if( ppvObject == NULL )
	{
		hr = E_INVALIDARG;
	}
	else
	{
		hr = CComObjectRoot::InternalQueryInterface(pThis, pEntries, iid, ppvObject);
	}

	return hr;
}

//IAccessor
HRESULT STDMETHODCALLTYPE CInterBaseCommand::AddRefAccessor( 
		HACCESSOR	   hAccessor,
		ULONG		  *pcRefCount)
{
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	CSmartLock	   sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::ReleaseIbRow()");

	if (pIbAccessor->pCreater != (void *)this)
		return DB_E_BADACCESSORHANDLE;

	pIbAccessor->ulRefCount++;
	if (pcRefCount)
		*pcRefCount = pIbAccessor->ulRefCount;

	return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::CreateAccessor( 
		DBACCESSORFLAGS	   dwAccessorFlags,
		ULONG			   cBindings,
		const DBBINDING	   rgBindings[],
		ULONG			   cbRowSize,
		HACCESSOR		  *phAccessor,
		DBBINDSTATUS	   rgStatus[])
{
	CSmartLock	   sm(&m_cs);
	HRESULT	hr;
	PIBACCESSOR	   pIbAccessor;
	IBCOLUMNMAP	  *pColumnMap;

	DbgFnTrace("CInterBaseCommand::CreateAccessor()");

	BeginMethod(IID_IAccessor);

	if( dwAccessorFlags & DBACCESSOR_PASSBYREF )
		return DB_E_BYREFACCESSORNOTSUPPORTED;

	if( ( dwAccessorFlags & ~( DBACCESSOR_ROWDATA | DBACCESSOR_PARAMETERDATA | DBACCESSOR_OPTIMIZED ) ) != 0 ||
	    ( dwAccessorFlags &  ( DBACCESSOR_ROWDATA | DBACCESSOR_PARAMETERDATA ) ) == 0 )
	{
#ifdef INFORMATIVEERRORS
		WCHAR wszError[64];

		swprintf( wszError, L"Invalid accessor flags: 0x%x", dwAccessorFlags );
		return m_pInterBaseError->PostError( wszError, 0, DB_E_BADACCESSORFLAGS );
#else
		return DB_E_BADACCESSORFLAGS;
#endif
	}

	if( ( dwAccessorFlags & DBACCESSOR_PARAMETERDATA ) != 0 )
		pColumnMap = m_pIbParamMap; 
	else
		pColumnMap = m_pIbColumnMap;

	if( phAccessor == NULL )
		return E_INVALIDARG;

	*phAccessor = NULL;

	hr = CreateIbAccessor(
			dwAccessorFlags,
			cBindings,
			rgBindings,
			cbRowSize,
			&pIbAccessor,
			rgStatus,
			0,
			m_pIbColumnMap);
	if (SUCCEEDED(hr))
	{
		pIbAccessor->pCreater = (void *)this;
		pIbAccessor->ulSequence = m_cIbAccessorCount++;
		pIbAccessor->pNext = m_pIbAccessorsToRelease;
		m_pIbAccessorsToRelease = pIbAccessor;
		*phAccessor = _EncodeHandle( pIbAccessor );
	}
	return hr;
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetBindings( 
		HACCESSOR		   hAccessor,
		DBACCESSORFLAGS	  *pdwAccessorFlags,
		ULONG			  *pcBindings,
		DBBINDING		 **prgBindings)
{
	CSmartLock	   sm(&m_cs);
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	DbgFnTrace("CInterBaseCommand::GetBindings()");

	if (pIbAccessor->pCreater != (void *)this)
		return DB_E_BADACCESSORHANDLE;

	return GetIbBindings(
			pIbAccessor,
			pdwAccessorFlags,
			pcBindings,
			prgBindings);
}

        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::ReleaseAccessor( 
		HACCESSOR	   hAccessor,
		ULONG		  *pcRefCount)
{
	CSmartLock	   sm(&m_cs);
	ULONG		   ulRefCount;
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	DbgFnTrace("CInterBaseCommand::ReleaseAccessor()");
	
	if (pIbAccessor->pCreater != (void *)this)
		return DB_E_BADACCESSORHANDLE;

	ReleaseIbAccessor( pIbAccessor, &ulRefCount, &m_pIbAccessorsToRelease );
	
	if (pcRefCount)
		*pcRefCount = ulRefCount;

	return S_OK;
}

#ifdef ICOLUMNSROWSET
//
// IColumnsRowset
//
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetAvailableColumns(
		ULONG	  *pcOptColumns,
		DBID	 **prgOptColumns)
{
	ULONG			   i;
	DBID			   rgOptColumns[] = {
		EXPAND_DBID( DBCOLUMN_BASETABLENAME ),
		EXPAND_DBID( DBCOLUMN_BASECOLUMNNAME ),
		EXPAND_DBID( DBCOLUMN_KEYCOLUMN ) };

	DbgFnTrace("CInterBaseRowset::GetAvailableColumns()");

	if( ! ( pcOptColumns && prgOptColumns ) )
		return E_INVALIDARG;

	if( ( *prgOptColumns = (DBID *)g_pIMalloc->Alloc( sizeof( rgOptColumns ) ) ) == NULL )
		return E_OUTOFMEMORY;

	for( i = 0; i < sizeof( rgOptColumns ) / sizeof( DBID ); i++ )
	{
		( *prgOptColumns )[i] = rgOptColumns[ i ];
	}

	*pcOptColumns = i;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetColumnsRowset(
		IUnknown	  *pUnkOuter,
		ULONG		   cOptColumns,
		const DBID	   rgOptColumns[],
		REFIID		   riid,
		ULONG		   cPropertySets,
		DBPROPSET	   rgPropertySets[],
		IUnknown	 **ppColRowset)
{
	CComPolyObject<CInterBaseRowset> *pColRowset = NULL;

	BEGIN_PREFETCHEDCOLUMNINFO( rgColumnInfo )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR,     SQL_VARYING,  DBCOLUMN_IDNAME,         L"DBCOLUMN_IDNAME" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,                38, DBTYPE_GUID,     SQL_TEXT + 1, DBCOLUMN_GUID,           L"DBCOLUMN_GUID" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,      sizeof(long), DBTYPE_UI4,      SQL_LONG + 1, DBCOLUMN_PROPID,         L"DBCOLUMN_PROPID" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR,     SQL_VARYING,  DBCOLUMN_NAME,           L"DBCOLUMN_NAME" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,      sizeof(long), DBTYPE_UI4,      SQL_LONG,     DBCOLUMN_NUMBER,         L"DBCOLUMN_NUMBER" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,     sizeof(short), DBTYPE_UI2,      SQL_SHORT,    DBCOLUMN_TYPE,           L"DBCOLUMN_TYPE" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,    sizeof(void *), DBTYPE_IUNKNOWN, SQL_LONG + 1, DBCOLUMN_TYPEINFO,       L"DBCOLUMN_TYPEINFO" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,      sizeof(long), DBTYPE_UI4,      SQL_LONG,     DBCOLUMN_COLUMNSIZE,     L"DBCOLUMN_COLUMNSIZE" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,     sizeof(short), DBTYPE_UI1,      SQL_SHORT,    DBCOLUMN_PRECISION,      L"DBCOLUMN_PRECISION" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,     sizeof(short), DBTYPE_UI1,      SQL_SHORT,    DBCOLUMN_SCALE,          L"DBCOLUMN_SCALE" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,      sizeof(long), DBTYPE_UI4,      SQL_LONG,     DBCOLUMN_FLAGS,          L"DBCOLUMN_FLAGS" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR,     SQL_VARYING,  DBCOLUMN_BASETABLENAME,  L"DBCOLUMN_BASETABLENAME" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0, MAX_SQL_NAME_SIZE, DBTYPE_WSTR,     SQL_VARYING,  DBCOLUMN_BASECOLUMNNAME, L"DBCOLUMN_BASECOLUMNNAME" )
		PREFETCHEDCOLUMNINFO_ENTRY2( 0,     sizeof(short), DBTYPE_BOOL,     SQL_SHORT,    DBCOLUMN_KEYCOLUMN,      L"DBCOLUMN_KEYCOLUMN" )
	END_PREFETCHEDCOLUMNINFO

	DbgFnTrace("CInterBaseCommand::GetColumnsRowset()");
	
	HRESULT    hr;
	ULONG      cbBuffer = 0;
	ULONG      cColumns = 11;
	ULONG      i;
	IBROW      *pIbRow;
	short      sValue;

	// Make sure that any optional columns are the ones we support.
	for (i = cColumns; i < PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo ); i++)
	{
		ULONG j;

		for (j = 0; j < cOptColumns; j++)
		{
			if (rgColumnInfo[i].columnid.uName.pwszName == rgOptColumns[j].uName.pwszName)
				break;
		}
		if (i == cOptColumns)
			return DB_E_BADCOLUMNID;
	}

	if( cOptColumns > 0 )
	{
		cColumns = PREFECTHEDCOLUMNINFOCOUNT( rgColumnInfo );

		if( ! m_fHaveKeyColumns &&
			FAILED( hr = GetKeyColumnInfo(this->m_pOutSqlda, this->m_pIbColumnMap, this->m_piscDbHandle, this->m_piscTransHandle ) ) )
		{
			return hr;
		}
		m_fHaveKeyColumns = true;
	}

	// Create the rowset object
	if( ( pColRowset = new CComPolyObject<CInterBaseRowset>( pUnkOuter ) ) == NULL )
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}
		
	if (FAILED(hr = pColRowset->m_contained.Initialize(
			this, 
			m_pOwningSession, 
			NULL,
			NULL,
			NULL,
			0,
			NULL,
			NULL,
			NULL,
			NULL,
			0,
			m_fUserCommand )))
		goto cleanup;

	pColRowset->m_contained.SetPrefetchedColumnInfo( rgColumnInfo, cColumns, NULL );

	if( FAILED( hr = pColRowset->QueryInterface( riid, (void **) ppColRowset ) ) )
	{
		goto cleanup;
	}

	//Populate ColumnsRowset Rowset
	for (i = 0; i < m_cColumns; i++)
	{
		pIbRow = pColRowset->m_contained.AllocRow( );

		//DBCOLUMN_IDNAME
		pColRowset->m_contained.SetPrefetchedColumnData( 0, pIbRow,   SQL_TEXT, m_pOutSqlda->sqlvar[i].sqlname_length, m_pOutSqlda->sqlvar[i].sqlname);

		//DBCOLUMN_GUID
		pColRowset->m_contained.SetPrefetchedColumnData( 1, pIbRow, IB_DEFAULT, IB_NULL, NULL );
					
		//DBCOLUMN_PROPID
		pColRowset->m_contained.SetPrefetchedColumnData( 2, pIbRow, IB_DEFAULT, IB_NULL, NULL );

		//DBCOLUMN_NAME
		pColRowset->m_contained.SetPrefetchedColumnData( 3, pIbRow,   SQL_TEXT, m_pOutSqlda->sqlvar[i].aliasname_length, m_pOutSqlda->sqlvar[i].aliasname);

		//DBCOLUMN_NUMBER
		pColRowset->m_contained.SetPrefetchedColumnData( 4, pIbRow, IB_DEFAULT, 0, (void *)&m_pIbColumnMap[i].iOrdinal );

		//DBCOLUMN_TYPE
		pColRowset->m_contained.SetPrefetchedColumnData( 5, pIbRow, IB_DEFAULT, 0, (void *)&m_pIbColumnMap[i].wType );

		//DBCOLUMN_TYPEINFO
		pColRowset->m_contained.SetPrefetchedColumnData( 6, pIbRow, IB_DEFAULT, IB_NULL, NULL );

		//DBCOLUMN_COLUMNSIZE
		pColRowset->m_contained.SetPrefetchedColumnData( 7, pIbRow, IB_DEFAULT, 0, (void *)&m_pIbColumnMap[i].ulColumnSize );

		//DBCOLUMN_PRECISION
		pColRowset->m_contained.SetPrefetchedColumnData( 8, pIbRow, IB_DEFAULT, 0, (void *)&(sValue = m_pIbColumnMap[i].bPrecision ) );

		//DBCOLUMN_SCALE
		pColRowset->m_contained.SetPrefetchedColumnData( 9, pIbRow, IB_DEFAULT, 0, (void *)&(sValue = m_pIbColumnMap[i].bScale ) );
		
		//DBCOLUMN_FLAGS
		pColRowset->m_contained.SetPrefetchedColumnData( 10, pIbRow, IB_DEFAULT, 0, (void *)&m_pIbColumnMap[i].dwFlags );

		if( cColumns > 10 )
		{
			//DBCOLUMN_BASETABLENAME
			pColRowset->m_contained.SetPrefetchedColumnData( 11, pIbRow, SQL_TEXT, m_pOutSqlda->sqlvar[i].relname_length, m_pOutSqlda->sqlvar[i].relname );
	
			//DBCOLUMN_BASECOLUMNNAME
			pColRowset->m_contained.SetPrefetchedColumnData( 12, pIbRow, SQL_TEXT, m_pOutSqlda->sqlvar[i].sqlname_length, m_pOutSqlda->sqlvar[i].sqlname);

			//DBCOLUMN_KEYCOLUMN
			pColRowset->m_contained.SetPrefetchedColumnData( 13, pIbRow, IB_DEFAULT, 0, (void *)&( sValue = ( ( m_pIbColumnMap[i].dwFlags & DBCOLUMNFLAGS_KEYCOLUMN ) != 0 ) ) );
		}
		pColRowset->m_contained.InsertPrefetchedRow(pIbRow);
	}

	pColRowset->m_contained.RestartPosition( NULL );
	m_cOpenRowsets++;

cleanup:
	if( FAILED( hr ) &&
		pColRowset != NULL )
	{
		delete pColRowset;
	}

	return hr;
}
#endif //ICOLUMNSROWSET

//
// ICommand
//
HRESULT STDMETHODCALLTYPE CInterBaseCommand::Cancel(void)
{
	DbgFnTrace("CInterBaseCommand::Cancel()");

	if( m_fExecuting )
	{
		return DB_E_CANTCANCEL;
	}

	return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::Execute( 
		IUnknown	  *pUnkOuter,
		REFIID		   riid,
		DBPARAMS	  *pParams,
		LONG		  *pcRowsAffected,
		IUnknown	 **ppRowset)
{
	HRESULT			   hr = S_OK;
	CSmartLock		   sm(&m_cs);
	void			  *pvParamBuffer = NULL;
	LONG			   cRowsAffected = DB_COUNTUNAVAILABLE;
	PIBACCESSOR		   pIbParamAccessor = NULL;
	CComPolyObject<CInterBaseRowset> *pRowset = NULL;

	DbgFnTrace( "CInterBaseCommand::Execute()" );

	BeginMethod(IID_ICommand);

	if( m_eCommandState == CommandStateInitial )
		return DB_E_NOCOMMAND;

	if( ( m_eCommandState == CommandStateUnprepared ||
	      ( m_eCommandState == CommandStateExecuted &&
		    m_iscStmtHandle == NULL ) ) &&
		FAILED(hr = Prepare(0)))
	{
		return hr;
	}

	m_fExecuting = TRUE;

	if( pParams != NULL )
	{
		XSQLVAR			  *pSqlVar;
		XSQLVAR			  *pSqlVarEnd;
		DBBINDING		  *pBinding;
		void			  *pvValue;
		DBSTATUS		   wStatus;
		isc_blob_handle	   iscBlobHandle = NULL;
		ISC_QUAD		   iscBlobId = {NULL, NULL};

		pIbParamAccessor = _DecodeAccessorHandle(pParams->hAccessor);
		pBinding = pIbParamAccessor->rgBindings;

		if( ! m_pIbParamMap && 
			FAILED( hr = DescribeBindings() ) )
		{
			goto cleanup;
		}

		pvParamBuffer = m_pOwningSession->AllocPage( m_cbParamBuffer );
		SetSqldaPointers(m_pInSqlda, m_pIbParamMap, pvParamBuffer);

		for(pSqlVar = m_pInSqlda->sqlvar, pSqlVarEnd = pSqlVar + (m_pInSqlda->sqld); pSqlVar < pSqlVarEnd; pSqlVar++, pBinding++ )
		{
			if( pBinding->dwPart & DBPART_VALUE )
				pvValue = _GetValuePtr(*pBinding, pParams->pData);
			else
				pvValue = NULL;

			if( pBinding->dwPart & DBPART_STATUS )
				wStatus = _GetStatus(*pBinding, pParams->pData);
			else
				wStatus = DBSTATUS_S_OK;

			*pSqlVar->sqlind = 0;

			if( wStatus & DBSTATUS_S_ISNULL || pvValue == NULL )
			{
				// It's NULL, we can do this pretty easily
				*pSqlVar->sqlind = -1;
				pSqlVar->sqltype |= 1;
			}
			else if( BASETYPE( pSqlVar->sqltype ) == SQL_BLOB &&
					 m_iscStmtType != isc_info_sql_stmt_delete )
			{
				// It's a blob, those are kinda special.  Create the new
				// BLOb, stuff the data in it, an put the BlobId into the sqlda
				ULONG	   cbBytesRead = 0;
				ULONG	   cbTotalBytes = 0;
				BYTE	  *pValue = NULL;
				BYTE	  *pValueToFree = NULL;
				char	   rgchBuffer[512];	//gds__alloc()

				switch( pBinding->wType )
				{
				case DBTYPE_IUNKNOWN:
					break;

				case DBTYPE_BYTES:
				case DBTYPE_STR:
					// Try to determine the length.  If they didn't bind the length and the
					// cbMaxLen == -1 then there is nothing we can do, so well just insert an
					// empty BLOB.  ADO Server cursor binds length and ADO Client Cursor gives
					// us IUnknown (ISequentialStream).
					if( pBinding->dwPart & DBPART_LENGTH )
					{
						cbTotalBytes = _GetLength(*pBinding, pParams->pData);
						pValue = (BYTE *)pvValue;
					}
					else
					{
						cbTotalBytes = pBinding->cbMaxLen != -1 ? pBinding->cbMaxLen : 0;
					}
					break;

				case DBTYPE_VARIANT:
					GetStringFromVariant( (char **)&pValueToFree, 0, &cbTotalBytes, (VARIANT *)pvValue );
					pValue = pValueToFree;
					break;

				default:
					if( pBinding->dwPart & DBPART_STATUS )
					{
						_GetStatus(*pBinding, pParams->pData ) = DBSTATUS_E_CANTCONVERTVALUE;
					}
					*pSqlVar->sqlind = -1;
					_ASSERTE(false);
					continue;
				}

				isc_create_blob2(
					m_iscStatus,
					m_piscDbHandle,
					m_piscTransHandle,
					&iscBlobHandle,
					&iscBlobId,
					0,
					NULL);

				if( pValue == NULL )
				{
					do
					{
						(*(ISequentialStream **)pvValue)->Read(rgchBuffer, sizeof(rgchBuffer), &cbBytesRead);

						if (cbBytesRead > 0)
							isc_put_segment(
									m_iscStatus,
									&iscBlobHandle,
									(USHORT)cbBytesRead,
									rgchBuffer);
					} while (cbBytesRead >= sizeof(rgchBuffer));
				}
				else
				{
					while( cbTotalBytes > 0 )
					{
						cbBytesRead = cbTotalBytes;

						if( cbBytesRead > SHRT_MAX )
							cbBytesRead = SHRT_MAX;
							
						isc_put_segment(
								m_iscStatus,
								&iscBlobHandle,
								(USHORT)cbBytesRead,
								(char *)pValue );
						pValue += cbBytesRead;
						cbTotalBytes -= cbBytesRead;
					}

					if( pValueToFree != NULL )
					{
						delete[] pValueToFree;
						pValueToFree = NULL;
					}
				}

				isc_close_blob(m_iscStatus, &iscBlobHandle);
				memcpy(pSqlVar->sqldata, &iscBlobId, pSqlVar->sqllen);
			}
			else
			{
				DBSTATUS wConverted;

				wConverted = InterBaseFromOleDbType(
						pSqlVar->sqltype, 
						pSqlVar->sqlsubtype, 
						pSqlVar->sqllen,
						pSqlVar->sqlscale,
						(BYTE *)pSqlVar->sqldata,
						pSqlVar->sqlind,
						pBinding->wType,
						pBinding->cbMaxLen,
						pBinding->bPrecision,
						pBinding->bScale,
						pvValue,
						wStatus);

				if( FAILED( wConverted ) )
				{
					_GetStatus(*pBinding, pParams->pData ) = DBSTATUS_E_CANTCONVERTVALUE;
					*pSqlVar->sqlind = -1;
					_ASSERTE(false);
					continue;
				}
			}
		}
	}

	switch (m_iscStmtType)
	{
	case isc_info_sql_stmt_select:
	case isc_info_sql_stmt_select_for_upd:
		if (isc_dsql_execute(m_iscStatus, m_piscTransHandle, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda))
		{
			hr = m_pInterBaseError->PostError(m_iscStatus);
			goto cleanup;
		}
		break;

	case isc_info_sql_stmt_exec_procedure:
		{
			void  *pvData;

			if( ( pvData = m_pOwningSession->AllocPage( m_cbRowBuffer ) ) == NULL )
			{
				hr = E_OUTOFMEMORY;
				goto cleanup;
			}
			SetSqldaPointers( m_pOutSqlda, m_pIbColumnMap, pvData );

			if( isc_dsql_execute2( m_iscStatus, m_piscTransHandle, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda, m_pOutSqlda ) )
			{
				m_pOwningSession->FreePage( pvData );
				hr = m_pInterBaseError->PostError( m_iscStatus );
				goto cleanup;
			}

			if( pIbParamAccessor != NULL )
			{			
				// If there are any output parameters, then we'll want to get their values
				// and put them into the consumer's buffer before we release the row.
				if( pIbParamAccessor->dwParamIO & DBPARAMIO_OUTPUT )
				{
					DBBINDING	  *pDbBinding = pIbParamAccessor->rgBindings;
					DBBINDING	  *pDbBindingEnd = pDbBinding + pIbParamAccessor->cBindInfo;
					XSQLVAR		  *pSqlVar;
					int			   nOutParam = 0;

					for( ; pDbBinding < pDbBindingEnd; pDbBinding++ )
					{
						if( pDbBinding->eParamIO & DBPARAMIO_INPUT )
							continue;

						pSqlVar = &(m_pOutSqlda->sqlvar[nOutParam]);
						GetIbValue(pDbBinding, (BYTE *)pParams->pData, pSqlVar, (BYTE *)pSqlVar->sqldata, *pSqlVar->sqlind );
						nOutParam++;
					}
				}
			}

			m_pOwningSession->FreePage( pvData );
		}
		break;

	case isc_info_sql_stmt_insert:
	case isc_info_sql_stmt_update:
	case isc_info_sql_stmt_delete:
		if( isc_dsql_execute( m_iscStatus, m_piscTransHandle, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda ) )
		{
			hr = m_pInterBaseError->PostError( m_iscStatus );
			goto cleanup;
		}
		
		cRowsAffected = 0;	// Trigger actual Rows Affected to be computed
		break;

	case isc_info_sql_stmt_ddl:
	case isc_info_sql_stmt_set_generator:
		{
			isc_tr_handle	iscDDLTransHandle = NULL;

			if( isc_start_transaction( m_iscStatus, &iscDDLTransHandle, 1, m_piscDbHandle, 0, NULL ) )
			{
				hr = m_pInterBaseError->PostError( m_iscStatus );
				goto cleanup;
			}
	
			m_fExecuting = true;
			if( isc_dsql_execute_immediate( m_iscStatus, m_piscDbHandle, &iscDDLTransHandle, 0, m_pszCommand, g_dso_usDialect, NULL ) )
			{

				isc_rollback_transaction( m_iscStatus, &iscDDLTransHandle );
				hr = m_pInterBaseError->PostError( m_iscStatus );
				goto cleanup;
			}

			if( isc_commit_transaction( m_iscStatus, &iscDDLTransHandle ) )
			{
				hr = m_pInterBaseError->PostError( m_iscStatus );
				goto cleanup;
			}
		}
		break;

	case isc_info_sql_stmt_start_trans:
	case isc_info_sql_stmt_commit:
	case isc_info_sql_stmt_rollback:
	default:
		hr = E_FAIL;
		goto cleanup;
	}

	// if the requested interface is IID_NULL, or it is a non rowset procuding command, we're done
	if( riid == IID_NULL || 
		! ( m_iscStmtType == isc_info_sql_stmt_select ||
		    m_iscStmtType == isc_info_sql_stmt_select_for_upd ) )
	{
		goto cleanup;
	}

	if( SUCCEEDED( hr = CComPolyObject<CInterBaseRowset>::CreateInstance( pUnkOuter, &pRowset ) ) )
	{
		pRowset->m_contained.Initialize(
				this, 
				m_pOwningSession, 
				m_piscDbHandle,
				m_piscTransHandle,
				m_iscStmtHandle,
				m_cbRowBuffer,
				m_pOutSqlda,
				m_pIbColumnMap,
				m_pInSqlda,
				m_pIbParamMap,
				m_cIbAccessorCount,
				m_fUserCommand );

		m_eCommandState = CommandStateExecuted;
		m_cOpenRowsets++;
		m_iscStmtHandle = NULL;

		if( FAILED( hr = pRowset->QueryInterface( riid, (void **)ppRowset ) ) )
			delete pRowset;
	}

cleanup:
	// This is here to trigger a commit or rollback if the transaction is auto-commit.  
	//m_pOwningSession->ReleaseTransHandle( true, SUCCEEDED(hr) );
	m_pOwningSession->AutoCommit( SUCCEEDED( hr ) );

	if( pvParamBuffer )
		m_pOwningSession->FreePage( pvParamBuffer );

	if( pcRowsAffected  )
	{
		if( cRowsAffected == 0 )
		{
			char	   rgchItems[] = {isc_info_sql_records, isc_info_end};
			char	   rgchResults[64];
			char	  *pch;

			if( ! isc_dsql_sql_info( m_iscStatus, &m_iscStmtHandle, sizeof(rgchItems), rgchItems, sizeof(rgchResults), rgchResults ) )
			{
				char	   chType;
				long	   lCount;
				short	   sLen;

				for( pch = rgchResults + 3; *pch != isc_info_end; )
				{
					chType = *pch++;
					sLen = (short)isc_vax_integer(pch, sizeof(short));
					pch += sizeof(short);
					lCount = isc_vax_integer(pch, sLen);
					pch += sLen;

					switch (chType)
					{
					case isc_info_req_update_count:
						cRowsAffected += lCount;
						break;
					case isc_info_req_delete_count:
						cRowsAffected += lCount;
						break;
					case isc_info_req_insert_count:
						cRowsAffected += lCount;
						break;
					}
				}
			}
		}

		*pcRowsAffected = cRowsAffected;
	}

	return hr;
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetDBSession( 
		REFIID		   riid,
		IUnknown	 **ppSession)
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::GetDBSession()");

	return m_pOwningSession->QueryInterface(riid, (void **)ppSession);
}

//ICommandPrepare
HRESULT STDMETHODCALLTYPE CInterBaseCommand::Prepare(
		ULONG	   cExpectedRuns)
{
	CSmartLock		   sm(&m_cs);
	HRESULT			   hr;
	isc_tr_handle	  *piscTransHandle = NULL;
	char	   rgchItems[] = {isc_info_sql_stmt_type};
	char	   rgchClump[8];

	DbgFnTrace( "CInterBaseCommand::Prepare()" );

	BeginMethod( IID_ICommandPrepare );

	if( m_eCommandState == CommandStateInitial )
		return DB_E_NOCOMMAND;

	if( m_cOpenRowsets )
		return DB_E_OBJECTOPEN;

	// Acquire the active transaction handle from the Session object.  If the current
	// transaction is auto-commit then this increments the transaction interest count, 
	// and will cause the auto-commit operation to be a commit retaining.
	if( FAILED( hr = m_pOwningSession->GetTransHandle( &piscTransHandle ) ) )
		return hr;

	if (!m_pOutSqlda)
	{
		if (!(m_pOutSqlda = (XSQLDA *)new char[XSQLDA_LENGTH(DEFAULT_COLUMN_COUNT)]))
		{
			hr = E_OUTOFMEMORY;
			goto cleanup;
		}

		m_pOutSqlda->sqln = DEFAULT_COLUMN_COUNT;
		m_pOutSqlda->version = SQLDA_VERSION1;
	}

	if( ( ( m_iscStmtHandle == NULL ) &&
		  isc_dsql_allocate_statement( m_iscStatus, m_piscDbHandle, &m_iscStmtHandle ) ) ||
		isc_dsql_prepare( m_iscStatus, piscTransHandle, &m_iscStmtHandle, 0, m_pszCommand, g_dso_usDialect, m_pOutSqlda ) ||
		isc_dsql_sql_info( m_iscStatus, &m_iscStmtHandle, sizeof(rgchItems), rgchItems, sizeof(rgchClump), rgchClump ) )
	{
		hr = m_pInterBaseError->PostError(m_iscStatus);
		goto cleanup;
	}

	_ASSERT( rgchClump[0] == isc_info_sql_stmt_type );
	m_iscStmtType = isc_vax_integer( rgchClump + 1 + sizeof( short ), (short)isc_vax_integer( rgchClump + 1, sizeof( short ) ) );

	if( m_pOutSqlda->sqld > m_pOutSqlda->sqln )
	{
		short	   n = m_pOutSqlda->sqld;
		
		delete[] m_pOutSqlda;

		if( ! ( m_pOutSqlda = (XSQLDA *)new char[XSQLDA_LENGTH(n)] ) )
		{
			hr = E_OUTOFMEMORY;
			goto cleanup;
		}
		
		m_pOutSqlda->sqln = n;
		m_pOutSqlda->version = SQLDA_VERSION1;
		
		if( isc_dsql_describe(m_iscStatus, &m_iscStmtHandle, g_dso_usDialect, m_pOutSqlda ) )
		{
			hr = m_pInterBaseError->PostError( m_iscStatus );
			goto cleanup;
		}
	}

	if( FAILED( hr = MakeColumnMap( m_pOutSqlda, &m_pIbColumnMap, &m_cbRowBuffer ) ) )
		goto cleanup;

	m_fHaveKeyColumns = false;
	m_cColumns = m_pOutSqlda->sqld;
	m_piscTransHandle = piscTransHandle;
	m_eCommandState = CommandStatePrepared;

cleanup:
	// If the Prepare failed, then we'll give the trnasaction back to the Session.  If the
	// transaction is auto-commit this will cause the interest count to decrease by one.
	if( FAILED( hr ) )
	{
		m_pOwningSession->ReleaseTransHandle( &m_piscTransHandle );
	}
	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseCommand::Unprepare(void)
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::Unprepare()");

	BeginMethod(IID_ICommandPrepare);

	if (m_cOpenRowsets)
		return DB_E_OBJECTOPEN;

	if (m_eCommandState == CommandStatePrepared ||
		m_eCommandState == CommandStateExecuted)
	{
		if( m_piscTransHandle )
		{
			// Give the transaction back to the Session since we're done with it.  It the
			// transaction is auto-commit this will cause the interest count to decrease by
			// one and commit the transaction.
			m_pOwningSession->ReleaseTransHandle( &m_piscTransHandle );
		}

		if (m_pIbColumnMap)
		{
			delete[] m_pIbColumnMap;
			m_pIbColumnMap = NULL;
		}

		if (m_pIbParamMap)
		{
			delete[] m_pIbParamMap;
			m_pIbParamMap = NULL;
		}

		if (m_pInSqlda)
			m_pInSqlda->sqld = 0;

		if (m_pOutSqlda)
			m_pOutSqlda->sqld = 0;

		if (isc_dsql_free_statement(m_iscStatus, &m_iscStmtHandle, DSQL_drop))
		{
			return m_pInterBaseError->PostError(m_iscStatus);
		}
	}

	m_eCommandState = CommandStateUnprepared;
	
	return S_OK;
}

//ICommandProperties
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetProperties( 
		const ULONG			   cPropertyIDSets,
		const DBPROPIDSET	   rgPropertyIDSets[],
		ULONG				  *pcPropertySets,
		DBPROPSET			 **prgPropertySets)
{
	CSmartLock	   sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::GetProperties()");
	
	if (!cPropertyIDSets)
	{
		DBPROPIDSET	rgPropIDSets[] = {NULL, -1, EXPANDGUID(DBPROPSET_ROWSET)};
		ULONG	cPropIDSets = sizeof(rgPropIDSets)/sizeof(DBPROPIDSET);

		return GetIbProperties(
				cPropIDSets,
				rgPropIDSets,
				pcPropertySets,
				prgPropertySets,
				_getIbProperties);
	}

	return GetIbProperties(
			cPropertyIDSets,
			rgPropertyIDSets,
			pcPropertySets,
			prgPropertySets,
			_getIbProperties);
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::SetProperties( 
		ULONG		   cPropertySets,
		DBPROPSET	   rgPropertySets[])
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::SetProperties()");

	return SetIbProperties(cPropertySets, rgPropertySets, _getIbProperties);
}

//ICommandText
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetCommandText( 
		GUID		  *pguidDialect,
		LPOLESTR	  *ppwszCommand)
{
	CSmartLock	sm(&m_cs);
	ULONG	   ulSize;
	WCHAR	  *pwszCommand = NULL;

	DbgFnTrace("CInterBaseCommand::GetCommandText()");

	if (!m_pszCommand)
		return DB_E_NOCOMMAND;

	ulSize = MultiByteToWideChar(
			CP_ACP,
			MB_PRECOMPOSED,
			m_pszCommand,
			-1,
			pwszCommand,
			0);

	if (!(pwszCommand = (WCHAR *)g_pIMalloc->Alloc((ulSize + 1) * sizeof(WCHAR))))
		return E_OUTOFMEMORY;

	ulSize = MultiByteToWideChar(
			CP_ACP,
			MB_PRECOMPOSED,
			m_pszCommand,
			-1,
			pwszCommand,
			ulSize);

	pwszCommand[ulSize] = L'\0';
	*ppwszCommand = pwszCommand;

return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CInterBaseCommand::SetCommandText( 
		REFGUID		   rguidDialect,
		LPCOLESTR	   pwszCommand)
{
	CSmartLock sm(&m_cs);
	HRESULT	   hr;
	VARIANTARG vtCommand;
	LPSTR      pchCmd;
	bool       fInSingleQuote = false;
	bool       fInDoubleQuote = false;

	DbgFnTrace("CInterBaseCommand::SetCommandText()");

	if (m_cOpenRowsets)
		return DB_E_OBJECTOPEN;

	if (!IsEqualGUID(rguidDialect, DBGUID_DBSQL))
		return DB_E_DIALECTNOTSUPPORTED;

	if( m_pszCommand != NULL )
	{
		delete[] m_pszCommand;
		m_pszCommand = NULL;
	}

	if ((m_eCommandState == CommandStatePrepared) ||
		(m_eCommandState == CommandStateExecuted))
		Unprepare();

	if (!pwszCommand || pwszCommand[0] == L'\0')
	{
		m_eCommandState = CommandStateInitial;

		return S_OK;
	}

	VariantInit(&vtCommand);
	vtCommand.vt = VT_BSTR;
	vtCommand.bstrVal = (USHORT *)pwszCommand;

	if( FAILED( hr = GetStringFromVariant( (char **)&m_pszCommand, 0, NULL, &vtCommand ) ) )
		return hr;

	for( pchCmd = m_pszCommand; *pchCmd != '\0'; pchCmd++ )
	{
		if( *pchCmd == '"' &&
			! fInSingleQuote )
		{
			fInDoubleQuote = !fInDoubleQuote;
		}
		else if( *pchCmd == '\'' &&
			! fInDoubleQuote )
		{
			fInSingleQuote = !fInSingleQuote;
		}
		else if( ! ( fInSingleQuote || fInDoubleQuote ) )
		{
			*pchCmd = toupper( *pchCmd );
		}
	}

	DbgTrace( m_pszCommand );
	DbgTrace( "\n" );
	m_eCommandState = CommandStateUnprepared;

	return S_OK;
}

//ICommandWithParameters
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetParameterInfo(
		ULONG		  *pcParams,
		DBPARAMINFO	 **prgParamInfo,
		OLECHAR		 **ppNamesBuffer)
{
	HRESULT	hr;
	ULONG		   iOrdinal;
	ULONG		   cParams = 0;
	DBPARAMINFO	  *pParamInfo;
	DBPARAMINFO	  *pParamInfoBuffer;

	DbgFnTrace("CInterBaseCommand::GetParameterInfo()");

	if( m_eCommandState == CommandStateInitial )
		return DB_E_NOCOMMAND;

	if( m_eCommandState == CommandStateUnprepared )
		return DB_E_NOTPREPARED;

	if( ! ( pcParams && prgParamInfo ) )
		return E_INVALIDARG;

	if( ! m_pIbParamMap && FAILED( hr = DescribeBindings() ) )
		return hr;

	cParams = m_cParams;

	if( m_iscStmtType == isc_info_sql_stmt_exec_procedure )
	{
		cParams += m_cColumns;
	}

	pParamInfoBuffer = (DBPARAMINFO *)g_pIMalloc->Alloc( cParams * sizeof( DBPARAMINFO ) );
	if( ! ( pParamInfo = pParamInfoBuffer ) )
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	// Get the information for the input parameters
	for( iOrdinal = 0 ; iOrdinal < m_cParams; iOrdinal++ )
	{
		pParamInfo->bPrecision = m_pIbParamMap[iOrdinal].bPrecision;
		pParamInfo->bScale = m_pIbParamMap[iOrdinal].bScale;
		pParamInfo->dwFlags = DBPARAMFLAGS_ISINPUT | ( pParamInfo->dwFlags & DBPARAMFLAGS_ISLONG );
		pParamInfo->iOrdinal = iOrdinal + 1;
		pParamInfo->pTypeInfo = NULL;
		pParamInfo->ulParamSize = m_pIbParamMap[iOrdinal].ulColumnSize;
		pParamInfo->wType = m_pIbParamMap[iOrdinal].wType;
		pParamInfo->pwszName = NULL;

		pParamInfo++;
	}

	// Get the information for the output parameters
	for( ; iOrdinal < cParams; iOrdinal++ )
	{
		pParamInfo->bPrecision = m_pIbColumnMap[iOrdinal - m_cParams].bPrecision;
		pParamInfo->bScale = m_pIbColumnMap[iOrdinal - m_cParams].bScale;
		pParamInfo->dwFlags = DBPARAMFLAGS_ISOUTPUT | ( pParamInfo->dwFlags & DBPARAMFLAGS_ISLONG );
		pParamInfo->iOrdinal = iOrdinal + 1;
		pParamInfo->pTypeInfo = NULL;
		pParamInfo->ulParamSize = m_pIbColumnMap[iOrdinal - m_cParams].ulColumnSize;
		pParamInfo->wType = m_pIbColumnMap[iOrdinal - m_cParams].wType;
		pParamInfo->pwszName = NULL;

		pParamInfo++;
	}

cleanup:
	*pcParams = pParamInfo - pParamInfoBuffer;

	if( pParamInfoBuffer == pParamInfo )
	{
		g_pIMalloc->Free( pParamInfoBuffer );
		*prgParamInfo = NULL;
	}
	else
	{
		*prgParamInfo = pParamInfoBuffer;
	}

	if( ppNamesBuffer )
	{
		*ppNamesBuffer = NULL;
	}

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseCommand::MapParameterNames(
		ULONG			   cParamNames,
		const OLECHAR	  *rgParamNames[],
		LONG			   rgParamOrdinals[])
{
	ULONG	iOrdinal;

	DbgFnTrace("CInterBaseCommand::MapParameterNames()");

	if( m_eCommandState == CommandStateInitial )
		return DB_E_NOCOMMAND;

	if( ! ( m_eCommandState == CommandStatePrepared ||
		    m_eCommandState == CommandStateExecuted ) )
		return DB_E_NOTPREPARED;

	if( cParamNames != 0 &&
		( rgParamNames == NULL || rgParamOrdinals == NULL ) )
		return E_INVALIDARG;

	for( iOrdinal = 0; iOrdinal < cParamNames; iOrdinal++ )
	{
		rgParamOrdinals[iOrdinal] = 0;
	}

return DB_E_ERRORSOCCURRED;
}

HRESULT STDMETHODCALLTYPE CInterBaseCommand::SetParameterInfo(
		ULONG					   cParams,
		const ULONG				   rgParamOrdinals[],
		const DBPARAMBINDINFO	   rgParamBindInfo[])
{
	HRESULT	   hr = S_OK;
	ULONG	   iOrdinal;
	short	   sInParams = 0;
	short	   sOutParams = 0;
	short	   sIbType;
	DBTYPE	   wType;
	int		   nSqlLen;
	DWORD	   dwHash;

	DbgFnTrace("CInterBaseCommand::SetParameterInfo()");

	if( cParams > 0 && rgParamOrdinals == NULL )
		return E_INVALIDARG;

	if( m_cOpenRowsets > 0 )
		return DB_E_OBJECTOPEN;

	if( cParams == 0 )
	{
		m_cParams = 0;
		if( m_pIbParamMap != NULL )
		{
			delete[] m_pIbParamMap;
			m_pIbParamMap = NULL;
		}
		return hr;
	}

	if( m_pIbParamMap != NULL )
	{
		hr = DB_S_TYPEINFOOVERRIDDEN;

		if( m_cParams != cParams )
		{
			if( m_pIbParamMap != NULL )
			{
				delete[] m_pIbParamMap;
				m_pIbParamMap = NULL;
			}
		}
	}

	if( m_pIbParamMap == NULL )
	{
		if( ( m_pIbParamMap = new IBCOLUMNMAP[ cParams ] ) == NULL )
			return E_OUTOFMEMORY;
	}

	if( m_pInSqlda == NULL )
	{
		if( ( m_pInSqlda = (XSQLDA *)new char[XSQLDA_LENGTH( cParams )] ) == NULL )
			return E_OUTOFMEMORY;

		m_pInSqlda->sqln = (short)cParams;
		m_pInSqlda->version = SQLDA_VERSION1;
	}

	m_cParams = cParams;
	m_cbParamBuffer = 0;

	for( iOrdinal = 0; iOrdinal < cParams; iOrdinal++ )
	{
		ULONG   ulParamSize = rgParamBindInfo[iOrdinal].ulParamSize;
		BYTE    bPrecision = rgParamBindInfo[iOrdinal].bPrecision;
		BYTE    bScale = rgParamBindInfo[iOrdinal].bScale;
		bool    fIsLong = false;

		m_pIbParamMap[iOrdinal].bPrecision = bPrecision;
		m_pIbParamMap[iOrdinal].bScale = bScale;

		if( rgParamOrdinals[iOrdinal] != iOrdinal + 1 )
		{
			hr = DB_E_BADORDINAL;
			break;
		}

		nSqlLen = ulParamSize;
		m_pIbParamMap[ iOrdinal ].iOrdinal = iOrdinal + 1;
		m_pIbParamMap[ iOrdinal ].dwFlags = rgParamBindInfo[iOrdinal].dwFlags;
		m_pIbParamMap[ iOrdinal ].wszAlias[ 0 ] = L'\0';
		m_pIbParamMap[ iOrdinal ].wszName[ 0 ] = L'\0';

		_ASSERTE( ( dwHash = HashString( L"DBTYPE_I2" ) )			== 0x005b8c86 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_I4" ) )			== 0x005b8c88 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_I8" ) )			== 0x005b8c8c );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_R4" ) )			== 0x005b8cac );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_R8" ) )			== 0x005b8cb0 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_DATE" ) )			== 0x05b8c9a5 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_DBDATE" ) )		== 0x5b8c98a5 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_DBTIME" ) )		== 0x5b8c9d09 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_DBTIMESTAMP" ) )	== 0x32749194 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_NUMERIC" ) )		== 0x6e3358c7 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_CHAR" ) )			== 0x05b8c996 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_VARCHAR" ) )		== 0x6e338c96 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_VARBINARY" ) )	== 0xe338ca31 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_LONGVARCHAR" ) )	== 0x33228c96 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_LONGVARBINARY" ) )== 0x3228ca31 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_WCHAR" ) )        == 0x16e33996 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_WVARCHAR" ) )     == 0xb8cf8c96 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_WLONGVARCHAR" ) ) == 0xcf228c96 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_BSTR" ) )         == 0x05b8ca52 );
		_ASSERTE( ( dwHash = HashString( L"DBTYPE_VARIANT" ) )      == 0x6e338ddc );

		if( rgParamBindInfo[iOrdinal].pwszDataSourceType == NULL )
		{
			wType = m_pIbParamMap[ iOrdinal ].wType;
			fIsLong = ( ( m_pIbParamMap[ iOrdinal ].dwFlags & DBCOLUMNFLAGS_ISLONG ) != 0 );
		}
		else
		{
			dwHash = HashString(rgParamBindInfo[iOrdinal].pwszDataSourceType);
			wType = DBTYPE_EMPTY;

			switch( dwHash )
			{
			case 0x005b8c86: /* DBTYPE_I2 */
				wType = DBTYPE_I2;
				break;
			case 0x005b8c88: /* DBTYPE_I4 */
				wType = DBTYPE_I4;
				break;
			case 0x005b8c8c: /* DBTYPE_I8 */
				if( g_dso_usDialect == 3 )
				{
					wType = DBTYPE_I8;
				}
				break;
			case 0x005b8cac: /* DBTYPE_R4 */
			case 0x005b8cb0: /* DBTYPE_R8 */
				if( ulParamSize == 4 )
					wType = DBTYPE_R4;
				else
					wType = DBTYPE_R8;
				break;
			case 0x05b8c9a5: /* DBTYPE_DATE */
				wType = DBTYPE_DATE;
				break;
			case 0x5b8c98a5: /* DBTYPE_DBDATE */
				if( g_dso_usDialect == 3 )
				{
					wType = DBTYPE_DBDATE;
				}
				else
				{
					wType = DBTYPE_DBTIMESTAMP;
				}
				break;
			case 0x5b8c9d09: /* DBTYPE_DBTIME */
				if( g_dso_usDialect == 3 )
				{
					wType = DBTYPE_DBTIME;
				}
				else
				{
					wType = DBTYPE_DBTIMESTAMP;
				}
				break;
			case 0x32749194: /* DBTYPE_DBTIMESTAMP */
				wType = DBTYPE_DBTIMESTAMP;
				break;
			case 0x6e3358c7: /* DBTYPE_NUMERIC */
				wType = DBTYPE_NUMERIC;
				break;
			case 0x05b8c996: /* DBTYPE_CHAR */
				wType = DBTYPE_STR;
				break;
			case 0x16e33996: /* DBTYPE_WCHAR */
				wType = DBTYPE_WSTR;
				break;
			case 0x6e338c96: /* DBTYPE_VARCHAR */
				wType = DBTYPE_STR;
				fIsLong = ( nSqlLen == -1 );
				break;
			case 0x33228c96: /* DBTYPE_LONGVARCHAR */
				wType = DBTYPE_STR;
				fIsLong = true;
				break;
			case 0xb8cf8c96: /* DBTYPE_VARWCHAR */
				wType = DBTYPE_WSTR;
				fIsLong = ( nSqlLen == -1 );
				break;
			case 0xcf228c96: /* DBTYPE_LONGVARWCHAR */
				wType = DBTYPE_WSTR;
				fIsLong = true;
				break;
			case 0xe338ca31: /* DBTYPE_VARBINARY */
			case 0x3228ca31: /* DBTYPE_LONGVARBINARY */
				wType = DBTYPE_BYTES;
				fIsLong = true;
				break;
			case 0x05b8ca52: /* DBTYPE_BSTR */
				wType = DBTYPE_BSTR;
				break;
			}
		}

		switch( wType )
		{
		case DBTYPE_I2:
			nSqlLen = sizeof(short);
			sIbType = SQL_SHORT;
			break;
		case DBTYPE_I4:
			nSqlLen = sizeof(long);
			sIbType = SQL_LONG;
			break;
		case DBTYPE_I8:
			nSqlLen = sizeof(LONGLONG);
			sIbType = SQL_INT64;
			break;
		case DBTYPE_R4:
			nSqlLen = sizeof(float);
			sIbType = SQL_FLOAT;
			break;
		case DBTYPE_R8:
			nSqlLen = sizeof(double);
			sIbType = SQL_DOUBLE;
			break;
		case DBTYPE_DATE:
			sIbType = SQL_TIMESTAMP; /* SQL_DATE */
			nSqlLen = sizeof(ISC_TIMESTAMP);
			break;
		case DBTYPE_DBDATE:
			sIbType = SQL_TYPE_DATE;
			nSqlLen = sizeof(ISC_DATE);
			break;
		case DBTYPE_DBTIME:
			sIbType = SQL_TYPE_TIME;
			nSqlLen = sizeof(ISC_TIME);
			break;
		case DBTYPE_DBTIMESTAMP:
			sIbType = SQL_TIMESTAMP; /* SQL_DATE */
			nSqlLen = sizeof(ISC_TIMESTAMP);
			break;
		case DBTYPE_NUMERIC:
			if( bPrecision <= 4 )
			{
				nSqlLen = sizeof(short);
				sIbType = SQL_SHORT;
			}
			else if( bPrecision <= 9 )
			{
				nSqlLen = sizeof(long);
				sIbType = SQL_LONG;
			}
			else if( g_dso_usDialect == 3 )
			{
				nSqlLen = sizeof(LONGLONG);
				sIbType = SQL_INT64;
			}
			else
			{
				nSqlLen = sizeof(double);
				sIbType = SQL_DOUBLE;
			}
			break;

		case DBTYPE_STR:
			if( ! fIsLong )
			{
				sIbType = SQL_VARYING;
				nSqlLen = nSqlLen + sizeof(short);
			}
			else
			{
				sIbType = SQL_BLOB;
				nSqlLen = sizeof(ISC_QUAD);
				ulParamSize = sizeof(IUnknown *);
				m_pInSqlda->sqlvar[sInParams].sqlsubtype = isc_blob_text;
			}
			break;
		case DBTYPE_WSTR:
			if( ! fIsLong )
			{
				sIbType = SQL_VARYING;
				nSqlLen = nSqlLen + sizeof(short);
			}
			else
			{
				sIbType = SQL_BLOB;
				nSqlLen = sizeof(ISC_QUAD);
				ulParamSize = sizeof(IUnknown *);
				m_pInSqlda->sqlvar[sInParams].sqlsubtype = isc_blob_text;
			}
			break;
		case DBTYPE_BYTES:
			sIbType = SQL_BLOB;
			nSqlLen = sizeof(ISC_QUAD);
			if( dwHash == 0x3228ca31 )
				ulParamSize = sizeof(IUnknown *);
			m_pInSqlda->sqlvar[sInParams].sqlsubtype = isc_blob_untyped;
			break;
		case DBTYPE_BSTR:
			sIbType = SQL_VARYING;
			nSqlLen = nSqlLen + sizeof(short);
			break;
		default :
			hr = DB_E_BADTYPENAME;
			break;
		}

		if( FAILED( hr ) )
		{
			break;
		}

		m_pIbParamMap[iOrdinal].wType = wType;
		m_pIbParamMap[iOrdinal].obData = m_cbParamBuffer;
		m_pIbParamMap[iOrdinal].ulColumnSize = ulParamSize;

		m_cbParamBuffer += ROUNDUP( ulParamSize, sizeof(int) );

		m_pIbParamMap[iOrdinal].obStatus = m_cbParamBuffer;

		m_cbParamBuffer += sizeof(DBSTATUS);

		m_pInSqlda->sqlvar[sInParams].sqlscale = -bScale;
		m_pInSqlda->sqlvar[sInParams].sqltype = sIbType;
		m_pInSqlda->sqlvar[sInParams].sqllen = (short)nSqlLen;
		m_pInSqlda->sqlvar[sInParams].sqlind = 0;
		m_pInSqlda->sqlvar[sInParams].sqlname_length = 0;
		m_pInSqlda->sqlvar[sInParams].sqlname[0] = '\0';

		if( m_pIbParamMap[iOrdinal].dwFlags & DBPARAMFLAGS_ISINPUT )
		{
			sInParams++;

			if( sOutParams > 0 )
			{
				hr = DB_E_BADORDINAL;
				break;
			}
		}
		if( m_pIbParamMap[iOrdinal].dwFlags & DBPARAMFLAGS_ISOUTPUT )
		{
			sOutParams++;
		}
	}

	if( SUCCEEDED( hr ) )
	{
		m_pInSqlda->sqld = sInParams;
	}
	else
	{
		m_cParams = 0;
		if( m_pIbParamMap != NULL )
		{
			delete[] m_pIbParamMap;
			m_pIbParamMap = NULL;
		}
		if( m_pInSqlda != NULL )
		{
			delete[] m_pInSqlda;
			m_pInSqlda = NULL;
		}
	}

	return hr;
}

//IConvertType
HRESULT STDMETHODCALLTYPE CInterBaseCommand::CanConvert( 
		DBTYPE			   wFromType,
		DBTYPE			   wToType,
		DBCONVERTFLAGS	   dwConvertFlags)
{
	DbgFnTrace("CInterBaseCommand::CanConvert()");

	return g_pIDataConvert->CanConvert(wFromType, wToType);
}

//IColumnsInfo
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetColumnInfo(
		ULONG		  *pcColumns, 
		DBCOLUMNINFO **prgInfo, 
		OLECHAR		 **ppStringsBuffer)
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::GetColumnInfo()");

	BeginMethod(IID_IColumnsInfo);

	if (!(m_eCommandState == CommandStatePrepared ||
		  m_eCommandState == CommandStateExecuted))
		return DB_E_NOTPREPARED;

	return GetIbColumnInfo(
			pcColumns,
			prgInfo,
			ppStringsBuffer,
			m_cColumns,
			m_pOutSqlda,
			m_pIbColumnMap);
}
 
HRESULT STDMETHODCALLTYPE CInterBaseCommand::MapColumnIDs(
		ULONG		   cColumnIDs, 
		const DBID	   rgColumnIDs[], 
		ULONG		   rgColumns[])
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseCommand::MapColumnIDs()");

	BeginMethod(IID_IColumnsInfo);

	if (m_eCommandState == CommandStateUnprepared)
		return DB_E_NOTPREPARED;

	return MapIbColumnIDs(
			cColumnIDs,
			rgColumnIDs,
			rgColumns,
			m_cColumns,
			m_pIbColumnMap);
}

//
// ISupportErrorInfo methods
//
STDMETHODIMP CInterBaseCommand::InterfaceSupportsErrorInfo
		(
		REFIID	iid
		)
{
	CSmartLock sm(&m_cs);

	if (IsEqualGUID(iid, IID_IUnknown))
	{
		return S_FALSE;
	}
	else 
	{
		return S_OK;
	}
}
