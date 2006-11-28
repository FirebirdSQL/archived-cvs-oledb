/*
 * Firebird Open Source OLE DB driver
 *
 * ROWSET.CPP - Source file for OLE DB rowset methods.
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
#include "../inc/ibaccessor.h"
#include "../inc/schemarowset.h"
#include "../inc/rowset.h"
#include "../inc/command.h"
#include "../inc/session.h"
#include "../inc/dso.h"
#include "../inc/smartlock.h"
#include "../inc/util.h"
#include "../inc/convert.h"
#include "../inc/blob.h"

CInterBaseRowset::CInterBaseRowset() :
	m_iscStmtHandle(NULL),
	m_cbRowBuffer( 0 ),
	m_pOwningSession(NULL),
	m_pOwningCommand(NULL),
	m_pInterBaseError(NULL),
	m_piscDbHandle(NULL),
	m_pInSqlda(NULL),
	m_pOutSqlda(NULL),
	m_pIbColumnMap(NULL),
	m_pFirstPrefetchedIbRow(NULL),
	m_pCurrentPrefetchedIbRow(NULL),
	m_pFreeIbRows(NULL),
	m_pMostRecentIbRow( NULL ),
	m_pIbAccessorsToRelease(NULL),
	m_pPagesToFree( NULL ),
	m_pPagePos( NULL ),
	m_pPageEnd( NULL ),
	m_fPrefetchedRowset( false )
{
	DbgFnTrace("CInterBaseRowset::CInterBaseRowset()");

	InitializeCriticalSection(&m_cs);
	m_cRows = 0;
	m_fEndOfRowset = false;
}

CInterBaseRowset::~CInterBaseRowset()
{
	IBACCESSOR	  *pIbAccessor;
	DbgFnTrace("CInterBaseRowset::~CInterBaseRowset()");

	while( m_pPagesToFree != NULL )
	{
		BYTE *pbPage = (BYTE *)m_pPagesToFree;
		m_pPagesToFree = ( (void **)pbPage )[ 0 ];

		m_pOwningSession->FreePage( (void *)pbPage );
	}

	while (m_pIbAccessorsToRelease)
	{
		pIbAccessor = m_pIbAccessorsToRelease;
		m_pIbAccessorsToRelease = pIbAccessor->pNext;
		delete[] pIbAccessor;
	}

	if (m_pOwningCommand)
	{
		m_pOwningCommand->DeleteRowset(m_iscStmtHandle);
		m_pOwningCommand->Release();
	}

	if (m_pOwningSession)
	{
		//m_pOwningSession->ReleaseTransHandle(FALSE);
		m_pOwningSession->Release();
	}

	if( m_pInterBaseError )
		delete m_pInterBaseError;

	DeleteCriticalSection(&m_cs);
}

HRESULT CInterBaseRowset::Initialize(
		CInterBaseCommand	  *pOwningCommand, 
		CInterBaseSession	  *pOwningSession, 
		isc_db_handle		  *piscDbHandle,
		isc_tr_handle		  *piscTransHandle,
		isc_stmt_handle		   iscStmtHandle,
		ULONG                  cbRowBuffer,
		XSQLDA				  *pOutSqlda,
		IBCOLUMNMAP			  *pIbColumnMap,
		XSQLDA				  *pInSqlda,
		IBCOLUMNMAP			  *pIbParamMap,
		ULONG				   ulAccessorSequence,
		bool				   fUserCommand )
{
	DbgFnTrace("CInterBaseRowset::Initialize()");

	m_fUserCommand = fUserCommand;
	m_cbRowBuffer = cbRowBuffer;
	m_pInSqlda = pInSqlda;
	m_pOutSqlda = pOutSqlda;
	m_cColumns = m_pOutSqlda == NULL ? 0 : m_pOutSqlda->sqld;
	m_pIbColumnMap = pIbColumnMap;
	m_iscStmtHandle = iscStmtHandle;
	m_piscTransHandle = piscTransHandle;
	m_piscDbHandle = piscDbHandle;
	m_pOwningSession = pOwningSession;
	m_pOwningCommand = pOwningCommand;
	m_ulAccessorSequence = ulAccessorSequence;
	if( m_pOwningCommand != NULL )
	{
		ULONG       cProperties;
		DBPROPSET   *pProperties = NULL;
		DBPROPID    rgPropID[] = { DBPROP_IColumnsRowset };
		DBPROPIDSET rgPropIDSet[1] = { rgPropID, sizeof(rgPropID)/sizeof(DBPROPID),EXPANDGUID(DBPROPSET_ROWSET) };
		
		m_pOwningCommand->AddRef();
		m_pOwningCommand->GetProperties( 1, rgPropIDSet, &cProperties, &pProperties );
		m_fIColumnsRowset   = ( V_BOOL( &pProperties[ 0 ].rgProperties[ 0 ].vValue) == VARIANT_TRUE );
		g_pIMalloc->Free( pProperties[ 0 ].rgProperties );
	}
	if( m_pOwningSession != NULL )
		m_pOwningSession->AddRef();

	return S_OK;
}

void *CInterBaseRowset::AllocPage( 
	ULONG  *pulSize )
{
	void *pvBuffer;

	if( ( pvBuffer = m_pOwningSession->AllocPage( *pulSize + sizeof( void * ) ) ) == NULL )
	{
		return NULL;
	}

	( (void **)pvBuffer )[ 0 ] = m_pPagesToFree;
	m_pPagesToFree = pvBuffer;
	*pulSize = ( (ULONG *)pvBuffer )[ -1 ];

	return ( (BYTE *)pvBuffer + sizeof( void * ) );
}

PIBROW CInterBaseRowset::AllocRow( void )
{
	PIBROW pIbRow = NULL;
	void *pvBuffer;
	
	if( m_pFreeIbRows != NULL )
	{
		pIbRow = m_pFreeIbRows;
		m_pFreeIbRows = pIbRow->pNext;
	}
	else
	{
		ULONG   ulReqSize = m_cbRowBuffer + sizeof( IBROW );
		if( (ULONG)( m_pPageEnd - m_pPagePos ) < ulReqSize )
		{
			if( ( pvBuffer = AllocPage( &ulReqSize ) ) == NULL )
			{
				return NULL;
			}

			m_pPagePos = (BYTE *)pvBuffer;
			m_pPageEnd = m_pPagePos + ulReqSize;
			m_pPagePos += sizeof( void * );
			_ASSERTE( m_pPagePos <= m_pPageEnd );
		}
		pvBuffer = m_pPagePos;
		m_pPagePos += ALIGN( m_cbRowBuffer + sizeof( IBROW ), sizeof( int ) );

		pIbRow = (IBROW *)pvBuffer;
		pIbRow->ulRefCount = 0;
		pIbRow->pNext = NULL;
	}

	if( pIbRow != m_pMostRecentIbRow )
	{
		SetSqldaPointers( m_pOutSqlda, m_pIbColumnMap, pIbRow->pbData );
		m_pMostRecentIbRow = pIbRow;
	}

	return pIbRow;
}

HRESULT CInterBaseRowset::InsertPrefetchedRow(
		PIBROW	   pIbRow)
{
	m_cRows++;
	pIbRow->pNext = NULL;
	pIbRow->ulRefCount = 1;
		
	if( m_pFirstPrefetchedIbRow )
	{
		m_pCurrentPrefetchedIbRow->pNext = pIbRow;
	}
	else
	{
		m_pFirstPrefetchedIbRow = pIbRow;
	}

	m_pCurrentPrefetchedIbRow = pIbRow;

	return S_OK;
}

HRESULT CInterBaseRowset::SetPrefetchedColumnData(
	ULONG    ulOrdinal,
	PIBROW   pIbRow,
	short    sqltype,
	short    ind,
	void    *pData )
{
	short  desttype = BASETYPE( m_pOutSqlda->sqlvar[ ulOrdinal ].sqltype );
	short  destlen  = m_pOutSqlda->sqlvar[ ulOrdinal ].sqllen;
	BYTE  *pbDestData = pIbRow->pbData + m_pIbColumnMap[ulOrdinal].obData;
	if( sqltype == IB_DEFAULT )
	{
		sqltype = desttype;
	}
	if( pData == NULL )
	{
		ind = IB_NULL;
	}
	
	if( ind != IB_NULL )
	{
		_ASSERTE( sizeof( pIbRow->pbData ) == sizeof( char ) );
		*(short *)( pIbRow->pbData + m_pIbColumnMap[ulOrdinal].obStatus ) = 0;
		if( BASETYPE( sqltype ) == desttype )
		{
			memcpy( pbDestData, pData, destlen );
		}
		else if( BASETYPE( sqltype ) == SQL_TEXT && desttype == SQL_VARYING )
		{
			if( ind == IB_NTS )
			{
				ind = strlen( (char *)pData );
			}
			else if( ind == IB_DEFAULT )
			{
				ind = destlen;
			}
			while( ind > 1 && ( (char *)pData )[ ind - 1] == ' ' )
				ind--;
			*(short *)pbDestData = ind;
			pbDestData += sizeof( short );
			memcpy( pbDestData, pData, ind );
			pbDestData[ ind ] = '\0';
		}
//		else if( BASETYPE( sqltype ) == SQL_VARYING && desttype == SQL_TEXT )
//		{
//		}
		else
		{
			_ASSERTE( "Unhandled type conversion" );
		}
	}
	else
	{
		*(short *)( pIbRow->pbData + m_pIbColumnMap[ulOrdinal].obStatus) = ind;
	}
	return S_OK;
}

HRESULT CInterBaseRowset::SetPrefetchedColumnInfo(
	const PrefetchedColumnInfo rgColumnInfo[],
	ULONG               cColumns,
	IBCOLUMNMAP  **ppIbColumnMap )
{
	ULONG          i;
	ULONG		   cbBuffer;
	XSQLVAR		  *pSqlVar;
	short          sColumnSize = 0;

	DbgFnTrace("CInterBaseRowset::SetPrefetchedColumnInfo()");

	cbBuffer = XSQLDA_LENGTH( cColumns );
	m_pOutSqlda = (XSQLDA *)AllocPage( &cbBuffer );
	m_pOutSqlda->sqln = (short)cColumns;
	m_pOutSqlda->sqld = (short)cColumns;
	cbBuffer = cColumns * sizeof( IBCOLUMNMAP );
	m_pIbColumnMap = (IBCOLUMNMAP *)AllocPage( &cbBuffer );

	pSqlVar = m_pOutSqlda->sqlvar;


	//Setup Required ColumnsRowset Columns
	for( i = 0, cbBuffer = 0; i < cColumns; i++ )
	{
		pSqlVar[i].aliasname[0] = '\0';
		pSqlVar[i].aliasname_length = 0;
		pSqlVar[i].ownname[0] = '\0';
		pSqlVar[i].ownname_length = 0;
		pSqlVar[i].relname[0] = '\0';
		pSqlVar[i].relname_length = 0;
		pSqlVar[i].sqldata = NULL;
		pSqlVar[i].sqlind = NULL;
		pSqlVar[i].sqlname[0] = '\0';
		pSqlVar[i].sqlname_length = 0;
		pSqlVar[i].sqlscale = 0;
		pSqlVar[i].sqlsubtype = 0;
		pSqlVar[i].sqltype = rgColumnInfo[i].sqltype;
		pSqlVar[i].sqldata = NULL;
		pSqlVar[i].sqlind = NULL;

		m_pIbColumnMap[ i ].dwFlags = ( rgColumnInfo[i].dwFlags | DBCOLUMNFLAGS_ISFIXEDLENGTH );
		if( NULLABLE( pSqlVar[i].sqltype ) )
			m_pIbColumnMap[ i ].dwFlags |= ( DBCOLUMNFLAGS_MAYBENULL | DBCOLUMNFLAGS_ISNULLABLE );
		
		switch( BASETYPE( rgColumnInfo[i].sqltype ) )
		{
		case SQL_DATE:
			sColumnSize = sizeof( ISC_QUAD );
			break;
		case SQL_TEXT:
			sColumnSize = rgColumnInfo[i].sqllen;
			break;
		case SQL_VARYING:
			m_pIbColumnMap[ i ].dwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
			sColumnSize = rgColumnInfo[i].sqllen + sizeof( short );
			break;
		case SQL_SHORT:
			sColumnSize = sizeof( short );
			break;
		case SQL_LONG:
			sColumnSize = sizeof( long );
			break;
		default:
			_ASSERT( false );
		}

		pSqlVar[i].sqllen = sColumnSize;
		m_pIbColumnMap[ i ].iOrdinal = i + 1;
		m_pIbColumnMap[ i ].cbLength = sColumnSize;
		wcscpy( m_pIbColumnMap[ i ].wszAlias, rgColumnInfo[i].pwszColumnName );
		wcscpy( m_pIbColumnMap[ i ].wszName, rgColumnInfo[i].pwszColumnName );
		m_pIbColumnMap[ i ].columnid = rgColumnInfo[i].columnid;
		m_pIbColumnMap[i].pvExtra = NULL;
		m_pIbColumnMap[i].ulRefCount = 0;
		m_pIbColumnMap[i].wType = rgColumnInfo[i].wType;
		m_pIbColumnMap[i].bScale = 0;
		m_pIbColumnMap[i].bPrecision = 0;
		m_pIbColumnMap[i].ulColumnSize = sColumnSize;

		m_pIbColumnMap[i].obData = cbBuffer;
		cbBuffer += ROUNDUP( sColumnSize, sizeof( ULONG ) );
		m_pIbColumnMap[i].obStatus = cbBuffer;
		cbBuffer += sizeof( short );
	}

	if( m_cbRowBuffer < cbBuffer )
	{
		m_cbRowBuffer = cbBuffer;
	}

	m_cColumns = cColumns;

	if( ppIbColumnMap )
	{
		*ppIbColumnMap = m_pIbColumnMap;
	}
	m_fPrefetchedRowset = true;

	return S_OK;
}

HRESULT CInterBaseRowset::FetchNextIbRows(
		HCHAPTER		   hChapter, 
		LONG			   lRowsOffset, 
		LONG			   cRows, 
		ULONG			  *pcRowsObtained, 
		PIBROW			  *prgpIbRows)
{
	CSmartLock	sm2(m_pOwningCommand->GetCriticalSection());
	HRESULT		   hr = S_OK;
	ISC_STATUS	   sc;
	LONG		   nRow;
	PIBROW		   pIbRow;
	PIBROW		  *ppIbRows = (PIBROW *)prgpIbRows;

	*pcRowsObtained = 0;

	for( nRow = 0; nRow < cRows; nRow++ )
	{
		if( m_fPrefetchedRowset )
		{
			if( m_pCurrentPrefetchedIbRow != NULL )
			{
				pIbRow = m_pCurrentPrefetchedIbRow;
				m_pCurrentPrefetchedIbRow = m_pCurrentPrefetchedIbRow->pNext;
				pIbRow->ulRefCount++;
			}
			else
			{
				hr = DB_S_ENDOFROWSET;
				break;
			}
		}
		else
		{
			if( ( pIbRow = AllocRow( ) ) != NULL )
			{
				pIbRow->ulRefCount = 1;

				if( ( sc = isc_dsql_fetch( m_iscStatus, &m_iscStmtHandle, 1, m_pOutSqlda ) ) == 0 )
				{
					m_cRows++;
				}
				else
				{
					const HROW rghRow[ 1 ] = { _EncodeHandle( pIbRow ) };

					this->ReleaseRows( 1, rghRow, 0, NULL, NULL );

					if( sc == 100L )
					{
						hr = DB_S_ENDOFROWSET;
					}
					else
					{
						hr = m_pInterBaseError->PostError( m_iscStatus );
					}
					break;
				}
			}
			else
			{
				hr = E_OUTOFMEMORY;
				break;
			}
		}

		ppIbRows[nRow] = pIbRow;
	}
	
	if (SUCCEEDED(hr))
		*pcRowsObtained = nRow;
	
	return hr;
}

HRESULT CInterBaseRowset::BeginMethod(
	REFIID	   riid)
{
	if (!(m_pInterBaseError ||
		 (m_pInterBaseError = new CInterBaseError)))
		return E_OUTOFMEMORY;

	m_pInterBaseError->ClearError(riid);

	return S_OK;
}

HRESULT WINAPI CInterBaseRowset::InternalQueryInterface(
	CInterBaseRowset *			pThis,
	const _ATL_INTMAP_ENTRY*	pEntries, 
	REFIID						iid, 
	void**						ppvObject)
{
	HRESULT    hr;
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::InternalQueryInterface()");

	DbgPrintGuid( iid );

	if( ppvObject == NULL )
	{
		hr = E_INVALIDARG;
	}
	else if( ( m_fPrefetchedRowset ||
	           ! m_fIColumnsRowset ) &&
		     IsEqualGUID( iid, IID_IColumnsRowset ) )
	{
		*ppvObject = NULL;
		hr = E_NOINTERFACE;
	} 
	else
	{
		hr = CComObjectRoot::InternalQueryInterface(pThis, pEntries, iid, ppvObject);
	}

	return hr;
}

//
// IAccessor
//
HRESULT STDMETHODCALLTYPE CInterBaseRowset::AddRefAccessor(
		HACCESSOR	   hAccessor, 
		ULONG		  *pcRefCount)
{
	CSmartLock sm(&m_cs);
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	DbgFnTrace("CInterBaseRowset::AddRefAccessor()");

	BeginMethod(IID_IAccessor);

	if (!((pIbAccessor->pCreater == (void *)this) ||
		  (pIbAccessor->pCreater == (void *)m_pOwningCommand &&
		   pIbAccessor->ulSequence <= m_ulAccessorSequence)))
		return DB_E_BADACCESSORHANDLE;

	pIbAccessor->ulRefCount++;
	if (pcRefCount)
		*pcRefCount = pIbAccessor->ulRefCount;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::CreateAccessor(
		DBACCESSORFLAGS	   dwAccessorFlags, 
		ULONG			   cBindings, 
		const DBBINDING	   rgBindings[], 
		ULONG			   cbRowSize, 
		HACCESSOR		  *phAccessor, 
		DBBINDSTATUS	   rgBindStatus[])
{
	CSmartLock sm(&m_cs);
	HRESULT	hr;
	PIBACCESSOR	   pIbAccessor = NULL;

	DbgFnTrace("CInterBaseRowset::CreateAccessor()");
	
	BeginMethod(IID_IAccessor);

	if( phAccessor == NULL )
		return E_INVALIDARG;

	if( dwAccessorFlags & DBACCESSOR_PASSBYREF )
		return DB_E_BYREFACCESSORNOTSUPPORTED;

	if( ( dwAccessorFlags & ~( DBACCESSOR_ROWDATA | DBACCESSOR_OPTIMIZED ) ) != 0 ||
	    ( dwAccessorFlags &  ( DBACCESSOR_ROWDATA ) ) == 0 )
	{
#ifdef INFORMATIVEERRORS
		WCHAR	wszError[64];

		swprintf( wszError, L"Invalid accessor flags: 0x%x", dwAccessorFlags );
		return m_pInterBaseError->PostError( wszError, 0, DB_E_BADACCESSORFLAGS );
#else
		return DB_E_BADACCESSORFLAGS;
#endif
	}

	hr = CreateIbAccessor(
			dwAccessorFlags,
			cBindings,
			rgBindings,
			cbRowSize,
			&pIbAccessor,
			rgBindStatus,
			m_cColumns,
			m_pIbColumnMap);
	if (SUCCEEDED(hr))
	{
		pIbAccessor->pCreater = (void *)this;
		pIbAccessor->ulSequence = 0;
		pIbAccessor->pNext = m_pIbAccessorsToRelease;
		m_pIbAccessorsToRelease = pIbAccessor;
		*phAccessor = _EncodeHandle( pIbAccessor );
	}
	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetBindings(
		HACCESSOR		   hAccessor,
		DBACCESSORFLAGS	  *pdwAccessorFlags,
		ULONG			  *pcBindings,
		DBBINDING		 **prgBindings)
{
	CSmartLock	   sm(&m_cs);
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	DbgFnTrace("CInterBaseRowset::GetBindings()");

	BeginMethod(IID_IAccessor);

	if (!((pIbAccessor->pCreater == (void *)this) ||
		  (pIbAccessor->pCreater == (void *)m_pOwningCommand &&
		   pIbAccessor->ulSequence <= m_ulAccessorSequence)))
		return DB_E_BADACCESSORHANDLE;

	return GetIbBindings(
			pIbAccessor,
			pdwAccessorFlags,
			pcBindings,
			prgBindings);
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::ReleaseAccessor(
		HACCESSOR	   hAccessor, 
		ULONG		  *pcRefCount)
{
	CSmartLock	   sm(&m_cs);
	ULONG		   ulRefCount;
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );

	DbgFnTrace("CInterBaseRowset::ReleaseAccessor()");

	BeginMethod(IID_IAccessor);

	if( hAccessor == NULL )
		return DB_E_BADACCESSORHANDLE;

	if( ( pIbAccessor->pCreater == (void *)m_pOwningCommand ) &&
		( pIbAccessor->ulSequence <= m_ulAccessorSequence ) )
		return m_pOwningCommand->ReleaseAccessor(hAccessor, pcRefCount);

	if( ! ( pIbAccessor->pCreater == (void *)this ) )
		return DB_E_BADACCESSORHANDLE;

	ReleaseIbAccessor( pIbAccessor, &ulRefCount, &m_pIbAccessorsToRelease );

	if( pcRefCount )
		*pcRefCount = ulRefCount;

	return S_OK;
}

//
//IColumnsInfo
//
HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetColumnInfo(
		ULONG		  *pcColumns, 
		DBCOLUMNINFO **prgInfo, 
		OLECHAR		 **ppStringsBuffer)
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::GetColumnInfo()");

	BeginMethod(IID_IColumnsInfo);

	return GetIbColumnInfo(
			pcColumns,
			prgInfo,
			ppStringsBuffer,
			m_cColumns,
			m_pOutSqlda,
			m_pIbColumnMap);
}
 
HRESULT STDMETHODCALLTYPE CInterBaseRowset::MapColumnIDs(
		ULONG		   cColumnIDs, 
		const DBID	   rgColumnIDs[], 
		ULONG		   rgColumns[])
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::MapColumnIDs()");

	BeginMethod(IID_IColumnsInfo);

	return MapIbColumnIDs(
			cColumnIDs,
			rgColumnIDs,
			rgColumns,
			m_cColumns,
			m_pIbColumnMap);
}

//
// IColumnsRowset
//
#ifdef ICOLUMNSROWSET
HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetAvailableColumns(
		ULONG	  *pcOptColumns,
		DBID	 **prgOptColumns)
{
	DbgFnTrace("CInterBaseRowset::GetAvailableColumns()");

	return m_pOwningCommand->GetAvailableColumns(
			pcOptColumns,
			prgOptColumns);
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetColumnsRowset(
		IUnknown	  *pUnkOuter,
		ULONG		   cOptColumns,
		const DBID	   rgOptColumns[],
		REFIID		   riid,
		ULONG		   cPropertySets,
		DBPROPSET	   rgPropertySets[],
		IUnknown	 **ppColRowset)
{
	DbgFnTrace("CInterBaseRowset::GetColumnsRowset()");

	return m_pOwningCommand->GetColumnsRowset(
			pUnkOuter,
			cOptColumns,
			rgOptColumns,
			riid,
			cPropertySets,
			rgPropertySets,
			ppColRowset);
}
#endif //ICOLUMNSROWSET

//
//IConvertType
//
HRESULT STDMETHODCALLTYPE CInterBaseRowset::CanConvert(
		DBTYPE			   wFromType, 
		DBTYPE			   wToType, 
		DBCONVERTFLAGS	   dwConvertFlags)
{
	DbgFnTrace("CInterBaseRowset::CanConvert()");

	return m_pOwningCommand->CanConvert(wFromType, wToType, dwConvertFlags);
}

//IRowset
HRESULT STDMETHODCALLTYPE CInterBaseRowset::AddRefRows(
		ULONG		   cRows, 
		const HROW	   rghRows[], 
		ULONG		   rgRefCounts[], 
		DBROWSTATUS	   rgRowStatus[])
{
	CSmartLock	sm(&m_cs);
	ULONG	   nRow;
	PIBROW	   pIbRow;

	DbgFnTrace("CInterBaseRowset::AddRefRows()");

	BeginMethod(IID_IRowset);

	for (nRow = 0; nRow < cRows; nRow++)
	{
		pIbRow = _DecodeRowHandle( rghRows[nRow] );
		pIbRow->ulRefCount++;
		
		if (rgRefCounts)
			rgRefCounts[nRow] = pIbRow->ulRefCount;
		if (rgRowStatus)
			rgRowStatus[nRow] = DBROWSTATUS_S_OK;
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetData(
		HROW		   hRow, 
		HACCESSOR	   hAccessor, 
		void		  *pData)
{
	CSmartLock	sm(&m_cs);
	ULONG		   iOrdinal;
	PIBROW		   pIbRow = _DecodeRowHandle( hRow );
	PIBACCESSOR	   pIbAccessor = _DecodeAccessorHandle( hAccessor );
	DBBINDING	  *pBinding = pIbAccessor->rgBindings;
	DBBINDING	  *pBindingEnd = pBinding + pIbAccessor->cBindInfo;
	DBSTATUS	   wStatus = DBSTATUS_S_OK;
	XSQLVAR		  *pSqlVar;
	BYTE		  *pSqlData;
	short		  *pSqlInd;

	DbgFnTrace("CInterBaseRowset::GetData()");

	BeginMethod(IID_IRowset);

	if (!((pIbAccessor->pCreater == (void *)this) ||
		  (pIbAccessor->pCreater == (void *)m_pOwningCommand &&
		   pIbAccessor->ulSequence <= m_ulAccessorSequence)))
	{
		return DB_E_BADACCESSORHANDLE;
	}

	for ( ; pBinding < pBindingEnd; pBinding++)
	{
		iOrdinal = pBinding->iOrdinal - 1;
		pSqlVar = (m_pOutSqlda->sqlvar) + iOrdinal;
		pSqlData = pIbRow->pbData + m_pIbColumnMap[iOrdinal].obData;
		pSqlInd = (short *)( pIbRow->pbData + m_pIbColumnMap[iOrdinal].obStatus );

		m_pOwningCommand->GetIbValue(pBinding, (BYTE *)pData, pSqlVar, pSqlData, *pSqlInd);
	}

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetNextRows(
		HCHAPTER	   hChapter, 
		LONG		   lRowsOffset, 
		LONG		   cRows, 
		ULONG		  *pcRowsObtained, 
		HROW		 **prghRows)
{
	CSmartLock	sm(&m_cs);
	HRESULT	   hr = S_OK;
	ULONG	   cRow;
	ULONG      cRowsObtained;
	PIBROW	  *rgpIbRows;

	DbgFnTrace("CInterBaseRowset::GetNextRows()");

	if( pcRowsObtained == NULL )
	{
		pcRowsObtained = &cRowsObtained;
	}

	BeginMethod(IID_IRowset);

	if (lRowsOffset < 0)
		return DB_E_CANTSCROLLBACKWARDS;
	if (cRows < 0)
		return DB_E_CANTFETCHBACKWARDS;
	
	if (m_fEndOfRowset)
	{
		*pcRowsObtained = 0;
		return DB_S_ENDOFROWSET;
	}
	
	rgpIbRows = (PIBROW *)alloca( cRows * sizeof(PIBROW) );

	if( *prghRows == NULL )
	{
		*prghRows = (HROW *)g_pIMalloc->Alloc(cRows * sizeof(HROW *));
		if (!prghRows)
			return E_OUTOFMEMORY;
	}

	hr = FetchNextIbRows(hChapter,
			lRowsOffset,
			cRows,
			pcRowsObtained,
			rgpIbRows);

	if (hr == DB_S_ENDOFROWSET)
		m_fEndOfRowset = TRUE;

	for (cRow = 0; cRow < *pcRowsObtained; cRow++)
	{
		(*prghRows)[cRow] = _EncodeHandle( rgpIbRows[cRow] );
	}

	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::ReleaseRows(
		ULONG			   cRows, 
		const HROW		   rghRows[], 
		DBROWOPTIONS	   rgRowOptions[],
		ULONG			   rgRefCounts[], 
		DBROWSTATUS		   rgRowStatus[])
{
	CSmartLock	sm(&m_cs);
	ULONG	   nRow;
	PIBROW	   pIbRow;
	DBROWSTATUS	dwRowStatus;

	DbgFnTrace("CInterBaseRowset::ReleaseRows()");

	BeginMethod(IID_IRowset);

	for (nRow = cRows; nRow > 0; nRow--)
	{
		pIbRow = _DecodeRowHandle( rghRows[nRow - 1] );
		if( --(pIbRow->ulRefCount) == 0 )
		{
			pIbRow->pNext = m_pFreeIbRows;
			m_pFreeIbRows = pIbRow;
			dwRowStatus = DBROWSTATUS_S_OK;
		}
		if (rgRefCounts)
			rgRefCounts[nRow - 1] = pIbRow->ulRefCount;
		if (rgRowStatus)
			rgRowStatus[nRow - 1] = dwRowStatus;
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::RestartPosition(HCHAPTER)
{
	HRESULT     hr = S_OK;
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::RestartPosition()");

	BeginMethod(IID_IRowset);

	if( m_cRows == 0 )
	{
		return S_OK;
	}

	if( ! m_fPrefetchedRowset )
	{
		if( isc_dsql_free_statement( m_iscStatus, &m_iscStmtHandle, DSQL_close ) ||
			isc_dsql_execute( m_iscStatus, m_piscTransHandle, &m_iscStmtHandle, g_dso_usDialect, m_pInSqlda ) )
		{
			hr = m_pInterBaseError->PostError(m_iscStatus);
		}
		else
		{
			m_cRows = 0;
		}
	}
	else
	{
		m_pCurrentPrefetchedIbRow = m_pFirstPrefetchedIbRow;
	}

	if( SUCCEEDED( hr ) )
	{
		m_fEndOfRowset = false;
	}

	return hr;
}

#ifdef IROWSETIDENTITY
/* IRowsetIdentity */
HRESULT STDMETHODCALLTYPE CInterBaseRowset::IsSameRow(
	HROW   hThisRow,
	HROW   hThatRow )
{
	DbgFnTrace("CInterBaseRowset::IsSameRow()");
	PIBROW   pThisIbRow = _DecodeRowHandle( hThisRow );
	PIBROW   pThatIbRow = _DecodeRowHandle( hThatRow );

	if( pThisIbRow == pThatIbRow )
	{
		return S_OK;
	}
	if( m_fPrefetchedRowset )
	{
		return S_FALSE;
	}
	return S_FALSE;
}
#endif /* IROWSETIDENTITY */

//IRowsetInfo
HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetProperties(
		const ULONG			   cPropertyIDSets, 
		const DBPROPIDSET	   rgPropertyIDSets[], 
		ULONG				  *pcPropertySets, 
		DBPROPSET			 **prgPropertySets)
{
	HRESULT hr;

	DbgFnTrace("CInterBaseRowset::GetProperties()");

	//Since almost all the rowset properties are Read-Only, I'll just delegate this 
	//to the commmand object
	hr = m_pOwningCommand->GetProperties(
			cPropertyIDSets, 
			rgPropertyIDSets, 
			pcPropertySets, 
			prgPropertySets);

	// If we succeeded then we need to fix-up the IColumnsRowset property value
	if( SUCCEEDED( hr ) )
	{
		DBPROPSET  *pPropertySet = *prgPropertySets;
		DBPROPSET  *pPropertySetEnd = pPropertySet + *pcPropertySets;
		DBPROP  *pProperty;
		DBPROP  *pPropertyEnd;

		for( ; pPropertySet < pPropertySetEnd; pPropertySet++ )
		{
			if( IsEqualGUID( pPropertySet->guidPropertySet, DBPROPSET_ROWSET ) )
			{
				pProperty = pPropertySet->rgProperties;
				pPropertyEnd = pProperty + pPropertySet->cProperties;

				for( ; pProperty < pPropertyEnd; pProperty++ )
				{
					switch( pProperty->dwPropertyID )
					{
					case DBPROP_IColumnsRowset:
						V_BOOL( &pProperty->vValue ) = m_fIColumnsRowset || ! m_fPrefetchedRowset ? VARIANT_TRUE : VARIANT_FALSE;
						break;
					}
				}
			}
		}
	}

	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetReferencedRowset(
		ULONG		   iOrdinal, 
		REFIID		   riid, 
		IUnknown	 **ppReferencedRowset)
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::GetReferencedRowset()");

	if (iOrdinal > m_cColumns)
		return DB_E_BADORDINAL;
	return DB_E_NOTAREFERENCECOLUMN;
}

HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetSpecification(
		REFIID		   riid, 
		IUnknown	 **ppSpecification )
{
	CSmartLock	sm(&m_cs);

	DbgFnTrace("CInterBaseRowset::GetSpecification()");

	if (m_fUserCommand)
		return m_pOwningCommand->QueryInterface(riid, (void **)ppSpecification);
	return m_pOwningSession->QueryInterface(riid, (void **)ppSpecification);
}

STDMETHODIMP CInterBaseRowset::InterfaceSupportsErrorInfo(
		REFIID	iid )
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
