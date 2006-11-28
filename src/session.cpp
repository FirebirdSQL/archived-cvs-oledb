/*
 * Firebird Open Source OLE DB driver
 *
 * SESSION.CPP - Source file for OLE DB session onject methods.
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
#include "../inc/rowset.h"
#include "../inc/command.h"
#include "../inc/session.h"
#include "../inc/dso.h"
#include "../inc/smartlock.h"
#include "../inc/schemarowset.h"

CInterBaseSession::CInterBaseSession() :
	m_pInterBaseError(NULL),
	m_pOwningDSO(NULL),
	m_piscDbHandle(NULL),
	m_iscTransHandle(NULL),
	m_ulTransLevel(0),
	m_pFreePages( NULL )
{
	DbgFnTrace("CInterBaseSession::CInterBaseSession()");

	InitializeCriticalSection(&m_cs);
	m_isoLevel = ISOLATIONLEVEL_UNSPECIFIED;
	m_cTransLocks = 0;
	m_fAutoCommit = TRUE;
}

CInterBaseSession::~CInterBaseSession()
{
	IBPROPERTY	  *pIbProperties;
	ULONG		   cIbProperties;

	DbgFnTrace("CInterBaseSession::~CInterBaseSession()");

	if ((pIbProperties = _getIbProperties(DBPROPSET_SESSION, &cIbProperties)) != NULL)
		ClearIbProperties(pIbProperties, cIbProperties);

	if( m_iscTransHandle )
	{
		if( m_fAutoCommit )
		{
			InternalCommit( false, XACTTC_SYNC );
		}
		else
		{
			InternalAbort( false );
		}
	}

	if (m_pOwningDSO)
	{
		m_pOwningDSO->DeleteSession();
		m_pOwningDSO->Release();
	}

	while( m_pFreePages != NULL )
	{
		BYTE *pbPage = (BYTE *)m_pFreePages;
		m_pFreePages = ( (void **)pbPage )[ 0 ];
		pbPage -= sizeof( ULONG );
		delete[] pbPage;
	}

	if( m_pInterBaseError )
		delete m_pInterBaseError;

	DeleteCriticalSection(&m_cs);
}

void *CInterBaseSession::AllocPage(
	ULONG   ulReqSize )
{
	CSmartLock	   sm(&m_cs);

	if( ulReqSize <= DEFAULT_PAGE_SIZE &&
		m_pFreePages != NULL )
	{
		void *pPage = m_pFreePages;
		m_pFreePages = ( (void **)pPage )[ 0 ];

		return pPage;
	}

	ULONG ulPageSize = MAX( ulReqSize, DEFAULT_PAGE_SIZE );
	BYTE *pbPage = new BYTE[ ulPageSize + sizeof( ULONG ) ];
	pbPage += sizeof( ULONG );
	( (ULONG *)pbPage )[ -1 ] = ulPageSize;
	return (void *)pbPage;
}

void CInterBaseSession::FreePage( void *pPage )
{
	CSmartLock	   sm(&m_cs);

	if( ( (ULONG *)pPage )[ -1 ] == DEFAULT_PAGE_SIZE )
	{
		( (void **)pPage )[ 0 ] = m_pFreePages;
		m_pFreePages = pPage;
	}
	else
	{
		delete (BYTE *)( (BYTE *)pPage - sizeof( ULONG ) );
	}
}

HRESULT CInterBaseSession::Initialize(
		CInterBaseDSO	  *pOwningDSO, 
		isc_db_handle	  *piscDbHandle)
{
	m_piscDbHandle = piscDbHandle;
	m_pOwningDSO = pOwningDSO;
	m_pOwningDSO->AddRef();

	return S_OK;
}

HRESULT CInterBaseSession::BeginMethod(
	REFIID	   riid)
{
	if (!(m_pInterBaseError ||
		 (m_pInterBaseError = new CInterBaseError)))
		return E_OUTOFMEMORY;

	m_pInterBaseError->ClearError(riid);

	return S_OK;
}

HRESULT CInterBaseSession::GetTransHandle(
	isc_tr_handle **ppiscTransHandle )
{
	CSmartLock	   sm(&m_cs);
	HRESULT		   hr = S_OK;
	bool		   fAutoCommit = m_fAutoCommit;

	DbgFnTrace( "CInterBaseSession::GetTransHandle()" );

	//
	// If we are in auto-commit mode and there is not an active transaction then
	// lookup the desired AutoCommitIsoLevel and start the transaction
	//
	if( m_fAutoCommit && !m_iscTransHandle )
	{
		ULONG		   cIbProperties;
		IBPROPERTY	  *pIbProperties;
		DWORD		   dwIsoLevel;

		pIbProperties = _getIbProperties(DBPROPSET_SESSION, &cIbProperties);
		dwIsoLevel = V_I4( GetIbPropertyValue( pIbProperties + cIbProperties - 1 ) );

		hr = StartTransaction( dwIsoLevel, 0, NULL, NULL );
	}

	if( !m_iscTransHandle )
		hr = XACT_E_NOTRANSACTION;

	//
	// The StartTransaction() call above set m_fAutoCommit to FALSE, we need to
	// restore it to its value when this function was called.
	//
	m_fAutoCommit = fAutoCommit;
	m_cTransLocks++;

	*ppiscTransHandle = &m_iscTransHandle;

	return hr;
}

HRESULT CInterBaseSession::AutoCommit( bool fSuccess )
{
	CSmartLock	   sm(&m_cs);

	if( ! m_fAutoCommit )
	{
		return S_OK;
	}

	if( fSuccess )
	{
		return InternalCommit( false, XACTTC_SYNC );
	}
	
	return InternalAbort( false );
}

HRESULT CInterBaseSession::ReleaseTransHandle(
	isc_tr_handle **ppiscTransHandle )
{
	CSmartLock	   sm(&m_cs);
	HRESULT		   hr = S_OK;

	DbgFnTrace("CInterBaseSession::ReleaseTransHandle()");

	*ppiscTransHandle = NULL;

	m_cTransLocks--;

	if( ! m_fAutoCommit )
	{
		return S_OK;
	}

	return InternalCommit( false, XACTTC_SYNC );
}

HRESULT CInterBaseSession::InternalAbort(
	bool    fRetaining )
{
	if (isc_rollback_transaction(m_iscStatus, &m_iscTransHandle))
		return E_UNEXPECTED;

	m_fAutoCommit = true;

	if (fRetaining)
	{
		return XACT_E_CANTRETAIN;
	}
	
	return S_OK;
}

HRESULT CInterBaseSession::InternalCommit(
	/* [in] */  bool    fRetaining,
	/* [in] */  DWORD   grfTC )
{
	if( grfTC & XACTTC_SYNC_PHASEONE )
	{
		isc_prepare_transaction( m_iscStatus, &m_iscTransHandle );
	}
	else if( grfTC == 0 || grfTC & (XACTTC_SYNC | XACTTC_SYNC_PHASETWO) )
	{
		if( m_cTransLocks > 0 )
		{
			isc_commit_retaining( m_iscStatus, &m_iscTransHandle );
		}
		else
		{
			isc_commit_transaction( m_iscStatus, &m_iscTransHandle );
		}
		m_fAutoCommit = true;
	}
	else
		return XACT_E_NOTSUPPORTED;


	return S_OK;
}

HRESULT WINAPI CInterBaseSession::InternalQueryInterface(
	CInterBaseSession *			pThis,
	const _ATL_INTMAP_ENTRY*	pEntries, 
	REFIID						iid, 
	void**						ppvObject)
{
	HRESULT    hr;
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseSession::InternalQueryInterface()");

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

//IGetDataSource
HRESULT STDMETHODCALLTYPE CInterBaseSession::GetDataSource(
		REFIID riid,
		IUnknown **ppDataSource)
{
	DbgFnTrace( "CInterBaseSession::GetDataSource()" );

	BeginMethod( IID_IGetDataSource );

	return m_pOwningDSO->QueryInterface(riid, (void **)ppDataSource);
}

//IOpenRowset
HRESULT STDMETHODCALLTYPE CInterBaseSession::OpenRowset(
		IUnknown *pUnkOuter,
		DBID *pTableID,
		DBID *pIndexID,
		REFIID riid,
		ULONG cProperties,
		DBPROPSET rgProperties[  ],
		IUnknown **ppRowset)
{
	CSmartLock sm(&m_cs);
	HRESULT	   hr;
	WCHAR	   wszCommand[] = L"select * from \0z12345678901234567890123456789012z";

	DbgFnTrace( "CInterBaseSession::OpenRowset()" );

	BeginMethod( IID_IOpenRowset );

	if( NULL != pIndexID )
		return DB_E_NOINDEX;

	if( NULL == pTableID )
		return E_INVALIDARG;

	if( DBKIND_NAME != pTableID->eKind )
		return DB_E_NOTABLE;

	if( NULL == pTableID->uName.pwszName )
		return DB_E_NOTABLE;

	if( cProperties && ( rgProperties == NULL ) )
		return E_INVALIDARG;

	// If Dialect 3 then quote the identifier name
	if( 3 == g_dso_usDialect )
		wcscat( wszCommand, L"\"" );

	wcsncat(wszCommand, pTableID->uName.pwszName, 32);

	if( 3 == g_dso_usDialect )
		wcscat(wszCommand, L"\"" );

	try
	{
		CComObject<CInterBaseCommand> *pCommand = NULL;

		if( NULL == ( pCommand = new CComObject<CInterBaseCommand> ) )
			return E_OUTOFMEMORY;

		if( FAILED( hr = pCommand->Initialize( this, m_piscDbHandle, false ) ) ||
			FAILED( hr = pCommand->SetProperties( cProperties, rgProperties) ) ||
			FAILED( hr = pCommand->SetCommandText( DBGUID_DBSQL, wszCommand ) ) ||
			FAILED( hr = pCommand->Execute(pUnkOuter, riid, NULL, NULL, ppRowset) ) )
			delete pCommand;

		return hr;
	}
	catch(...)
	{
		return E_FAIL;
	}
}

//ISessionProperties
HRESULT STDMETHODCALLTYPE CInterBaseSession::GetProperties(
		ULONG cPropertyIDSets,
		const DBPROPIDSET rgPropertyIDSets[],
		ULONG *pcPropertySets,
		DBPROPSET **prgPropertySets)
{
	CSmartLock	   sm(&m_cs);

	DbgFnTrace("CInterBaseSession::GetProperties()");
	
	if (!cPropertyIDSets)
	{
		DBPROPIDSET	rgPropIDSets[] = {{NULL, -1, EXPANDGUID(DBPROPSET_SESSION)}, {NULL, -1, EXPANDGUID(DBPROPSET_IBASE_SESSION)}};
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

HRESULT STDMETHODCALLTYPE CInterBaseSession::SetProperties(
		ULONG cPropertySets,
		DBPROPSET rgPropertySets[])
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseSession::SetProperties()");

	return SetIbProperties(cPropertySets, rgPropertySets, _getIbProperties);
}

//IDBCreateCommand [opt]
HRESULT STDMETHODCALLTYPE CInterBaseSession::CreateCommand(
		/* [in] */ IUnknown *pUnkOuter,
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ IUnknown **ppCommand)
{
	CSmartLock	   sm(&m_cs);
	HRESULT		   hr;
	CComPolyObject<CInterBaseCommand> *pCommand = NULL;

	DbgFnTrace("CInterBaseSession::CreateCommand()");

	BeginMethod(IID_IDBCreateCommand);

	if( FAILED( hr = CComPolyObject<CInterBaseCommand>::CreateInstance( pUnkOuter, &pCommand ) ) )
		return hr;
		
	if( SUCCEEDED( hr = pCommand->m_contained.Initialize( this, m_piscDbHandle, true ) ) )
		hr = pCommand->QueryInterface( riid, (void **) ppCommand );

	if( FAILED( hr ) )
		delete pCommand;
	
	return hr;
}

//IDBSchemaRowset
HRESULT STDMETHODCALLTYPE CInterBaseSession::GetRowset(
		IUnknown       *pUnkOuter,
		REFGUID         rguidSchema,
		ULONG           cRestrictions,
		const VARIANT   rgRestrictions[],
		REFIID          riid,
		ULONG           cPropertySets,
		DBPROPSET       rgPropertySets[],
		IUnknown      **ppRowset)
{
	HRESULT		hr = E_INVALIDARG;
	PREFETCHEDROWSETCALLBACK pfnRowsetCallback = NULL;
	CComPolyObject<CInterBaseRowset> *pRowset = NULL;
	CComObject<CInterBaseCommand> *pCommand = NULL;

	DbgFnTrace("CInterBaseSession::GetRowset()");

	DbgPrintGuid( rguidSchema );

	BeginMethod(IID_IDBSchemaRowset);

	if( ( cRestrictions > 0 &&
		  rgRestrictions == NULL ) ||
		( ppRowset == NULL ) )
	{
		return E_INVALIDARG;
	}

	if( rguidSchema == DBSCHEMA_COLUMNS )
	{
		pfnRowsetCallback = SchemaColumnsRowset;
	}
	else if( rguidSchema == DBSCHEMA_TABLES )
	{
		pfnRowsetCallback = SchemaTablesRowset;
	}
	else if( rguidSchema == DBSCHEMA_PROVIDER_TYPES )
	{
		pfnRowsetCallback = SchemaProviderTypesRowset;
	}
#ifdef OPTIONAL_SCHEMA_ROWSETS
	else if( rguidSchema == DBSCHEMA_FOREIGN_KEYS )
	{
		pfnRowsetCallback = SchemaForeignKeysRowset;
	}
	else if( rguidSchema == DBSCHEMA_INDEXES )
	{
		pfnRowsetCallback = SchemaIndexesRowset;
	}
	else if( rguidSchema == DBSCHEMA_PRIMARY_KEYS )
	{
		pfnRowsetCallback = SchemaPrimaryKeysRowset;
	}
	else if( rguidSchema == DBSCHEMA_PROCEDURES )
	{
		pfnRowsetCallback = SchemaProceduresRowset;
	}
#endif //OPTIONAL_SCHEMA_ROWSETS

	if( pfnRowsetCallback == NULL )
		return E_INVALIDARG;

	if( NULL == ( pCommand = new CComObject<CInterBaseCommand> ) ||
		NULL == ( pRowset = new CComPolyObject<CInterBaseRowset>( pUnkOuter ) ) )
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	if( FAILED( hr = pCommand->SetProperties( cPropertySets, rgPropertySets ) ) )
	{
		goto cleanup;
	}

	pRowset->m_contained.Initialize( pCommand, this, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, 0, false);

	if( SUCCEEDED( hr = pfnRowsetCallback( m_piscDbHandle, pRowset, cRestrictions, rgRestrictions ) ) )
	{
		hr = pRowset->QueryInterface( riid, (void **) ppRowset );
	}

cleanup:
	if( FAILED( hr ) )
	{
		if( pCommand != NULL )
			delete pCommand;
		if( pRowset != NULL )
			delete pRowset;	
	}

	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseSession::GetSchemas(
		ULONG    *pcSchemas,
		GUID    **prgSchemas,
		ULONG   **prgRestrictions)
{
	HRESULT	     hr = S_OK;
	static GUID  rgSchemas[] = {
#ifdef OPTIONAL_SCHEMA_ROWSETS
		EXPANDGUID( DBSCHEMA_FOREIGN_KEYS ),
		EXPANDGUID( DBSCHEMA_INDEXES ),
		EXPANDGUID( DBSCHEMA_PRIMARY_KEYS ),
		EXPANDGUID( DBSCHEMA_PROCEDURES ),
#endif //OPTIONAL_SCHEMA_ROWSETS
		EXPANDGUID( DBSCHEMA_COLUMNS ),
		EXPANDGUID( DBSCHEMA_TABLES ),
		EXPANDGUID( DBSCHEMA_PROVIDER_TYPES ) };
	static ULONG rgRestrictions[] = {
#ifdef OPTIONAL_SCHEMA_ROWSETS
		0x00000024,	/* Foreign Keys */
		0x0000001C,	/* Indexes */
		0x00000004,	/* Primary Keys */
		0x0000000C,	/* Procedures */
#endif //OPTIONAL_SCHEMA_ROWSETS
		0x0000000C,	/* Columns */
		0x0000000C,	/* Tables */
		0x00000003	/* Provider Types */ } ;

	DbgFnTrace("CInterBaseSession::GetSchemas()");

	_ASSERTE( ( sizeof( rgSchemas ) / sizeof( GUID ) ) == ( sizeof( rgRestrictions ) / sizeof( ULONG ) ) );

	BeginMethod(IID_IDBSchemaRowset);

	*prgSchemas = (GUID *)g_pIMalloc->Alloc( sizeof( rgSchemas ) );
	*prgRestrictions = (ULONG *)g_pIMalloc->Alloc( sizeof( rgRestrictions ) );

	if( *prgSchemas == NULL ||
		*prgRestrictions == NULL )
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	*pcSchemas = sizeof( rgSchemas ) / sizeof( GUID );
	memcpy( *prgSchemas, rgSchemas, sizeof( rgSchemas ) );
	memcpy( *prgRestrictions, rgRestrictions, sizeof( rgRestrictions ) );

cleanup:
	if( FAILED( hr ) )
	{
		if( *prgSchemas != NULL )
			g_pIMalloc->Free( *prgSchemas );
		if( *prgRestrictions != NULL )
			g_pIMalloc->Free( rgRestrictions );
		*pcSchemas = 0;
		*prgSchemas = NULL;
		*prgRestrictions = NULL;
	}

	return hr;
}

//
// ISupportErrorInfo [opt]
//
STDMETHODIMP CInterBaseSession::InterfaceSupportsErrorInfo
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

//
// ITransaction [opt]
//
HRESULT STDMETHODCALLTYPE CInterBaseSession::Abort(
		/* [in] */	BOID *	pboidReason,
		/* [in] */  BOOL	fRetaining,
		/* [in] */	BOOL	fAsync)
{
	CSmartLock	   sm(&m_cs);

	DbgFnTrace("CInterBaseSession::Abort()");

	BeginMethod(IID_ITransaction);

	if( m_iscTransHandle == NULL ||
		m_fAutoCommit == true )
	{
		return XACT_E_NOTRANSACTION;
	}

	if( fAsync == TRUE )
	{
		return XACT_E_NOTSUPPORTED;
	}

	return InternalAbort( ( fRetaining == TRUE ) );
}

HRESULT STDMETHODCALLTYPE CInterBaseSession::Commit(
		/* [in] */  BOOL	fRetaining,
		/* [in] */  DWORD   grfTC,
		/* [in] */	DWORD	grfRM)
{
	HRESULT       hr = S_OK;
	CSmartLock    sm(&m_cs);

	DbgFnTrace("CInterBaseSession::Commit()");

	BeginMethod(IID_ITransaction);

	if( m_iscTransHandle == NULL ||
		m_fAutoCommit == true )
	{
		return XACT_E_NOTRANSACTION;
	}

	hr = InternalCommit( ( fRetaining == TRUE ), grfTC );

	if( m_iscStatus[ 1 ] )
	{
		m_pInterBaseError->PostError( m_iscStatus );
		hr = XACT_E_COMMITFAILED;
	}

	return hr;
}

HRESULT STDMETHODCALLTYPE CInterBaseSession::GetTransactionInfo(
		/* [out] */ XACTTRANSINFO *	pInfo)
{
	DbgFnTrace("CInterBaseSession::GetTransactionInfo()");

	BeginMethod(IID_ITransaction);

	if( pInfo == NULL )
	{
		return E_INVALIDARG;
	}

	pInfo->uow = BOID_NULL;
	pInfo->isoLevel = m_isoLevel;
	pInfo->grfTCSupported = ISOLATIONLEVEL_READCOMMITTED | ISOLATIONLEVEL_CURSORSTABILITY | ISOLATIONLEVEL_REPEATABLEREAD;
	pInfo->isoFlags = 0;
	pInfo->grfRMSupported = 0;
	pInfo->grfTCSupportedRetaining = 0;
	pInfo->grfRMSupportedRetaining = 0;
	
	return S_OK;
}

//ITransactionLocal [opt]
HRESULT STDMETHODCALLTYPE CInterBaseSession::GetOptionsObject(
		/* [out] */	ITransactionOptions ** ppOptions)
{
	DbgFnTrace("CInterBaseSession::GetOptionsObject()");

	BeginMethod(IID_ITransactionLocal);

	return E_FAIL;
}
		
HRESULT STDMETHODCALLTYPE CInterBaseSession::StartTransaction(
		/* [in] */	ISOLEVEL				isoLevel,
		/* [in] */	ULONG					isoFlags,
		/* [in] */	ITransactionOptions *	pOtherOptions,
		/* [out] */ ULONG *					pulTransactionLevel)
{
	CSmartLock sm(&m_cs);
	HRESULT hr = S_OK;

	DbgFnTrace("CInterBaseSeccion::StartTransaction()");

	BeginMethod(IID_ITransactionLocal);

	if (isoFlags)
		return XACT_E_NOISORETAIN;

	if( m_iscTransHandle )
	{
		if( m_fAutoCommit )
		{
			// There is a current transaction, but it is auto-commit.  Check
			// to see that its isolation level is the same as, or an enforcement
			// of the requested transaction isolation level.  If not we fail.
			if( isoLevel != ISOLATIONLEVEL_UNSPECIFIED &&
				isoLevel > m_isoLevel )
			{
				hr = XACT_E_ISOLATIONLEVEL;
			}
		}
		else
		{
			// The current transaction is manual-commit, so this is an attempt
			// to start a nested transaction.  We cannot do that.
			hr = XACT_E_XTIONEXISTS;
		}
	}
	else
	{
		// There is no current transaction, so try to start one with the
		// requested isolation level.
		char	chMode;
		char	pchTPB[16];
		short	sLenTPB = 0;

		IBPROPERTY *pIbProperties = _getIbProperties(DBPROPSET_IBASE_SESSION, NULL);
		bool fWait = (V_BOOL(GetIbPropertyValue(pIbProperties + IBASE_SESS_WAIT_INDEX)) != 0);
		bool fRecordVersions = (V_BOOL(GetIbPropertyValue(pIbProperties + IBASE_SESS_RECORDVERSIONS_INDEX)) != 0);
		bool fReadOnly = (V_BOOL(GetIbPropertyValue(pIbProperties + IBASE_SESS_READONLY_INDEX)) != 0);

		switch (isoLevel)
		{
		case ISOLATIONLEVEL_UNSPECIFIED:
		case ISOLATIONLEVEL_BROWSE:	// ISOLATIONLEVEL_READUNCOMMITTED
		case ISOLATIONLEVEL_REPEATABLEREAD:
			m_isoLevel = ISOLATIONLEVEL_REPEATABLEREAD;
			chMode = isc_tpb_concurrency;
			break;

		case ISOLATIONLEVEL_READCOMMITTED:	// ISOLATIONLEVEL_CURSORSTABILITY
			m_isoLevel = ISOLATIONLEVEL_READCOMMITTED;
			chMode = isc_tpb_read_committed;
			break;

		case ISOLATIONLEVEL_SERIALIZABLE:	// ISOLATIONLEVEL_ISOLATED
			m_isoLevel = ISOLATIONLEVEL_SERIALIZABLE;
			chMode = isc_tpb_consistency;
			break;

//		case ISOLATIONLEVEL_CHAOS:
		default:
			return XACT_E_ISOLATIONLEVEL;
		}

		pchTPB[sLenTPB++] = isc_tpb_version3;
		pchTPB[sLenTPB++] = (fReadOnly ? isc_tpb_read : isc_tpb_write);
		pchTPB[sLenTPB++] = chMode;
		if( chMode == isc_tpb_read_committed )
			pchTPB[sLenTPB++] = (fRecordVersions ? isc_tpb_rec_version : isc_tpb_no_rec_version);
		pchTPB[sLenTPB++] = (fWait ? isc_tpb_wait : isc_tpb_nowait);
		if( isc_start_transaction( m_iscStatus, &m_iscTransHandle, 1, m_piscDbHandle, sLenTPB, pchTPB ) )
		{
			hr = m_pInterBaseError->PostError(m_iscStatus);
		}
	}
	
	if( SUCCEEDED( hr ) )
	{
		// Everything is happy, so turn auto-commit off and retuen the nesting level.
		m_fAutoCommit = false;

		if( pulTransactionLevel )
		{
			*pulTransactionLevel = 1;
		}
	}

	return hr;
}

