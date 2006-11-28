/*
 * Firebird Open Source OLE DB driver
 *
 * DSO.CPP - Source file for OLE DB Data SOurce Object methods.
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
//#define INITGUID
//#include <initguid.h>
#include <msdaguid.h>
//#include <msdadc.h>
#include "../inc/resource.h"
#include "../inc/iboledb.h"
#include "../inc/property.h"
#include "../inc/ibaccessor.h"
#include "../inc/command.h"
#include "../inc/session.h"
#include "../inc/dso.h"
#include "../inc/smartlock.h"
#include "../inc/convert.h"
#include "../inc/util.h"

IMalloc			*g_pIMalloc = NULL;
IDataConvert	*g_pIDataConvert = NULL;
BOOL CALLBACK DialogProc(HWND, UINT, WPARAM, LPARAM);

// InterBase Version 6.0 Specific Globals
int g_dso_usDialect = 1;
ISCENCODESQLDATE   g_dso_pIscEncodeSqlDate   = NULL;
ISCENCODESQLTIME   g_dso_pIscEncodeSqlTime   = NULL;
ISCENCODETIMESTAMP g_dso_pIscEncodeTimestamp = NULL;
ISCDECODESQLDATE   g_dso_pIscDecodeSqlDate   = NULL;
ISCDECODESQLTIME   g_dso_pIscDecodeSqlTime   = NULL;
ISCDECODETIMESTAMP g_dso_pIscDecodeTimestamp = NULL;

CInterBaseDSO::CInterBaseDSO()
{
	DbgFnTrace( "CInterBaseDSO::CInterBaseDSO()" );

	m_pInterBaseError = new CInterBaseError;

	m_dwStatus = DSF_UNINITIALIZED;

	// Internal implementation members
	InitializeCriticalSection( &m_cs );
	m_cOpenSessions = 0;

	// Database connection stuff
	m_pchDPB = NULL;
	m_cbDPB = 0;
	m_iscDbHandle = NULL;

	// Get an interface pointer for IMalloc
	if( g_pIMalloc == NULL )
		CoGetMalloc( MEMCTX_TASK, &g_pIMalloc );
	else g_pIMalloc->AddRef();

	// Get an interface pointer for IDataConvert
	if( g_pIDataConvert == NULL )
		CoCreateInstance( CLSID_OLEDB_CONVERSIONLIBRARY, NULL, CLSCTX_INPROC_SERVER, IID_IDataConvert, (void **)&g_pIDataConvert );
	else
		g_pIDataConvert->AddRef();
}

CInterBaseDSO::~CInterBaseDSO()
{
	IBPROPERTY	  *pIbProperties;
	ULONG		   cIbProperties;

	DbgFnTrace( "CInterBaseDSO::~CInterBaseDSO()" );

	if( m_pInterBaseError != NULL )
		delete m_pInterBaseError;

	if( m_iscDbHandle )
		isc_detach_database( NULL, &m_iscDbHandle );

	if( ( pIbProperties = _getIbProperties( DBPROPSET_DATASOURCE, &cIbProperties ) ) != NULL )
		ClearIbProperties( pIbProperties, cIbProperties );

	if( ( pIbProperties = _getIbProperties( DBPROPSET_DATASOURCEINFO, &cIbProperties ) ) != NULL )
		ClearIbProperties(pIbProperties, cIbProperties);

	if( ( pIbProperties = _getIbProperties( DBPROPSET_DBINIT, &cIbProperties ) ) != NULL )
		ClearIbProperties(pIbProperties, cIbProperties);

	if( g_pIMalloc )
		g_pIMalloc->Release();

	if( g_pIDataConvert )
		g_pIDataConvert->Release();

	DeleteCriticalSection( &m_cs );
}

HRESULT CInterBaseDSO::DeleteSession(void)
{
	CSmartLock sm( &m_cs );

	DbgFnTrace( "CInterBaseDSO::DeleteSession()" );

	m_cOpenSessions--;

	return S_OK;
}

HRESULT CInterBaseDSO::CreateSession(
	IUnknown	  *pUnkOuter,
	REFIID		   riid,
	IUnknown	 **ppDBSession)
{
	HRESULT		   hr;
	CSmartLock	   sm(&m_cs);
	CComPolyObject<CInterBaseSession> *pSession = NULL;

	DbgFnTrace( "CInterBaseDSO::CreateSession()" );

	if( ! ( m_dwStatus & DSF_INITIALIZED ) )
		return E_UNEXPECTED;

	// Create our session object
	hr = CComPolyObject<CInterBaseSession>::CreateInstance( pUnkOuter, &pSession );

	if( SUCCEEDED( hr ) )		
		hr = pSession->m_contained.Initialize( this, &m_iscDbHandle );

	if( SUCCEEDED( hr ) )
		hr = pSession->QueryInterface( riid, (void **) ppDBSession );
	
	if( SUCCEEDED( hr ) )
		m_cOpenSessions++;
	else
		delete pSession;

	return hr;
}

STDMETHODIMP CInterBaseDSO::GetClassID(
	CLSID	  *pClassID)
{
	DbgFnTrace( "CInterBaseDSO::GetClassID()" );

	if( ! pClassID )
		return E_INVALIDARG;

	*pClassID = CLSID_InterBaseProvider;

	return S_OK;
}

//
// IDBInfo
//
HRESULT STDMETHODCALLTYPE CInterBaseDSO::GetKeywords(
		LPOLESTR	  *ppwszKeywords)
{
	DbgFnTrace( "CInterBaseDSO::GetKeywords()" );

	//OLECHAR szKeywords[] = L"ACTIVE,AFTER,ASCENDING,AUTO,AUTOINC,BASE_NAME,BEFORE,BLOB,BOOLEAN,BYTES,CAST,CHECK_POINT_LENGTH,COMMITTED,COMPUTED,CONDITIONAL,CONTAINING,CREATE,CURRENT,DATABASE,DEBUG,DESCENDING,DO,ENTRY_POINT,EXIT,FILE,FILTER,FUNCTION,GDSCODE,GENERATOR,GEN_ID,GROUP_COMMIT_WAIT_TIME,IF,INACTIVE,INDEX,INPUT_TYPE,LONG,LENGTH,LOGFILE,LOG_BUFFER_SIZE,MANUAL,MAXIMUM_SEGMENT,MERGE,MESSAGE,MODULE_NAME,MONEY,NUM_LOG_BUFFERS,OUTPUT_TYPE,OVERFLOW,PAGE_SIZE,PAGE,PAGES,PARAMETER,PASSWORD,PLAN,POST_EVENT,PROTECTED,RAW_PARTITIONS,RDB$DB_KEY,RECORD_VERSION,RESERV,RESERVING,RETAIN,RETURNING_VALUES,RETURNS,SEGMENT,SHARED,SHADOW,SINGULAR,SNAPSHOT,SORT,STABILITY,STARTING,STARTS,STATISTICS,SUB_TYPE,SUSPEND,UNCOMMITTED,VARIABLE,WAIT,WHILE";

	if(ppwszKeywords == NULL )
		return E_INVALIDARG;

	if( ! ( m_dwStatus & DSF_INITIALIZED ) )
		return E_UNEXPECTED;

	//_ASSERTE((wcslen(szKeywords) + 1) * sizeof(OLECHAR) == sizeof(szKeywords));

	//*ppwszKeywords = (LPOLESTR)g_pIMalloc->Alloc( sizeof(szKeywords) );
	//if( *ppwszKeywords == NULL )
	//	return E_OUTOFMEMORY;

	//wcscpy( *ppwszKeywords, szKeywords );

	return E_FAIL;
}

HRESULT STDMETHODCALLTYPE CInterBaseDSO::GetLiteralInfo(
		ULONG			   cRequestedLiterals,
		const DBLITERAL	   rgRequestedLiterals[],
		ULONG			  *pcLiteralInfo,
		DBLITERALINFO	 **prgLiteralInfo,
		OLECHAR			 **ppCharBuffer)
{
	ULONG		   cchCharBuffer = 0;
	ULONG		   iLiteral;
	ULONG		   cLiterals = cRequestedLiterals;
	ULONG		   cLiteralsSet = 0;
	DBLITERAL	   rgSupportedLiterals[] = { DBLITERAL_QUOTE_PREFIX, DBLITERAL_QUOTE_SUFFIX, DBLITERAL_CHAR_LITERAL };
	DBLITERAL	  *pLiterals = (DBLITERAL *)rgRequestedLiterals;
	DBLITERALINFO *pLiteralInfo = NULL;

	DbgFnTrace( "CInterBaseDSO::GetLiteralInfo()" );

	if( ! ( m_dwStatus & DSF_INITIALIZED ) )
		return E_UNEXPECTED;

	if( cRequestedLiterals != 0 && rgRequestedLiterals == NULL ||
		pcLiteralInfo == NULL ||
		prgLiteralInfo == NULL ||
		ppCharBuffer == NULL )
		return E_INVALIDARG;

	if( cRequestedLiterals == 0 )
	{
		cLiterals = sizeof(rgSupportedLiterals) / sizeof(DBLITERAL);
		pLiterals = rgSupportedLiterals;
	}

	*pcLiteralInfo = 0;
	*ppCharBuffer = NULL;
	pLiteralInfo = (DBLITERALINFO *)g_pIMalloc->Alloc( cLiterals * sizeof(DBLITERALINFO) );
	if( (*prgLiteralInfo = pLiteralInfo) == NULL )
		return E_OUTOFMEMORY;

	for( iLiteral = 0; iLiteral < cLiterals; iLiteral++)
	{
		pLiteralInfo[iLiteral].lt = pLiterals[iLiteral];
		pLiteralInfo[iLiteral].pwszInvalidChars = NULL;
		pLiteralInfo[iLiteral].pwszInvalidStartingChars = NULL;

		switch( pLiterals[iLiteral] )
		{
		case DBLITERAL_QUOTE_PREFIX:
		case DBLITERAL_QUOTE_SUFFIX:
			if( g_dso_usDialect == 3 )
			{
				pLiteralInfo[iLiteral].fSupported = TRUE;
				pLiteralInfo[iLiteral].pwszLiteralValue = LPOLESTR(L"\"");
				pLiteralInfo[iLiteral].cchMaxLen = wcslen( pLiteralInfo[iLiteral].pwszLiteralValue );
				cchCharBuffer += pLiteralInfo[iLiteral].cchMaxLen + sizeof(WCHAR);
			}
			else
			{
				pLiteralInfo[iLiteral].fSupported = FALSE;
				pLiteralInfo[iLiteral].pwszLiteralValue = NULL;
				pLiteralInfo[iLiteral].cchMaxLen = 0;
			}
			break;

		case DBLITERAL_CHAR_LITERAL:
			pLiteralInfo[iLiteral].fSupported = TRUE;
			pLiteralInfo[iLiteral].pwszLiteralValue = LPOLESTR(L"'");
			pLiteralInfo[iLiteral].cchMaxLen = wcslen( pLiteralInfo[iLiteral].pwszLiteralValue );
			cchCharBuffer += pLiteralInfo[iLiteral].cchMaxLen + sizeof(WCHAR);
			break;

		default:
			pLiteralInfo[iLiteral].fSupported = FALSE;
			pLiteralInfo[iLiteral].pwszLiteralValue = NULL;
			pLiteralInfo[iLiteral].cchMaxLen = 0;
		}

		if(pLiteralInfo[iLiteral].fSupported)
		{
			cLiteralsSet++;
		}
	}

	if( cchCharBuffer != 0 )
	{
		int	cchString;
		OLECHAR	  *pCharBuffer;


		pCharBuffer = (OLECHAR *)g_pIMalloc->Alloc( cchCharBuffer * sizeof(OLECHAR) );
		if( (*ppCharBuffer = pCharBuffer) == NULL )
		{
			g_pIMalloc->Free(*prgLiteralInfo);
			*prgLiteralInfo = NULL;
			*pcLiteralInfo = 0;
			return E_OUTOFMEMORY;
		}
		for( iLiteral = 0; iLiteral < cLiterals; iLiteral++ )
		{
			if( pLiteralInfo[iLiteral].pwszLiteralValue != NULL )
			{
				cchString = wcslen(pLiteralInfo[iLiteral].pwszLiteralValue);
				wcscpy( pCharBuffer, pLiteralInfo[iLiteral].pwszLiteralValue );
				pLiteralInfo[iLiteral].pwszLiteralValue = pCharBuffer;
				pCharBuffer += cchString + 1;
			}
			if( pLiteralInfo[iLiteral].pwszInvalidStartingChars != NULL )
			{
				cchString = wcslen(pLiteralInfo[iLiteral].pwszInvalidStartingChars);
				wcscpy( pCharBuffer, pLiteralInfo[iLiteral].pwszInvalidStartingChars );
				pLiteralInfo[iLiteral].pwszInvalidStartingChars = pCharBuffer;
				pCharBuffer += cchString + 1;
			}
			if( pLiteralInfo[iLiteral].pwszInvalidChars != NULL )
			{
				cchString = wcslen(pLiteralInfo[iLiteral].pwszInvalidChars);
				wcscpy( pCharBuffer, pLiteralInfo[iLiteral].pwszInvalidChars );
				pLiteralInfo[iLiteral].pwszInvalidChars = pCharBuffer;
				pCharBuffer += cchString + 1;
			}
		}
	}

	*pcLiteralInfo = cLiterals;

	return cLiteralsSet == 0 ? DB_E_ERRORSOCCURRED : S_OK;
}

//
// IDBInitialize
//
HRESULT STDMETHODCALLTYPE CInterBaseDSO::Initialize(void)
{
	CSmartLock sm(&m_cs);
	ULONG		 cIbProperties;
	IBPROPERTY	*pIbProperties;

	BeginMethod(IID_IDBInitialize);

	DbgFnTrace("CInterBaseDSO::Initialize()");
	if (m_dwStatus & DSF_INITIALIZED)
		return DB_E_ALREADYINITIALIZED;

	ISC_STATUS		   iscStatus[20];
	char *pszValue = (char *)alloca( _MAX_PATH * 2 );
	DWORD	dwProp;

	dpbinit( &m_pchDPB, &m_cbDPB, isc_dpb_version1 );

	pIbProperties = _getIbProperties(DBPROPSET_DBINIT, &cIbProperties);

	dwProp = V_I4( GetIbPropertyValue( pIbProperties + INIT_PROMPT_INDEX ) );
	if (dwProp == DBPROMPT_PROMPT ||
		((dwProp == DBPROMPT_COMPLETE || dwProp == DBPROMPT_COMPLETEREQUIRED) &&
		 !( (pIbProperties + AUTH_USERID_INDEX)->fIsSet && (pIbProperties + AUTH_PASSWORD_INDEX)->fIsSet && (pIbProperties + INIT_DATASOURCE_INDEX)->fIsSet)))
	{
		dwProp = (DWORD)GetIbPropertyValue(pIbProperties + INIT_HWND_INDEX)->lVal;
		if( !DialogBoxParam(g_hInstance, MAKEINTRESOURCE(IDD_LOGIN), (HWND)dwProp, DialogProc, (LPARAM)pIbProperties) )
		{
			return DB_SEC_E_AUTH_FAILED;
		}
	}

	if( SUCCEEDED( GetStringFromVariant( (char **)&pszValue, _MAX_PATH, NULL, GetIbPropertyValue( pIbProperties + AUTH_USERID_INDEX ) ) ) )
	{
		dpbcat(&m_pchDPB, &m_cbDPB, isc_dpb_user_name, (void *)pszValue);
	}

	if( SUCCEEDED( GetStringFromVariant( (char **)&pszValue, _MAX_PATH, NULL, GetIbPropertyValue( pIbProperties + INIT_PROVIDERSTRING_INDEX ) ) ) )
	{
		char *pchStart;
		char *pchPos;
		char *pchEnd;

		pchPos = pchStart = pchEnd = pszValue;

		int	nKey = 0;

		while (true)
		{
			if (!iswspace(*pchPos))
				pchEnd++;

			if (pchStart == pchPos && iswspace(*pchPos))
			{
				pchStart++;
			}
			
			if (*pchPos == '=' && nKey == 0)
			{
				if (strnicmp(pchStart, "USER ROLE", pchEnd - pchStart) == 0)
				{
					nKey = 1;
				}
				else if (strnicmp(pchStart, "SQL DIALECT", pchEnd - pchStart) == 0)
				{
					nKey = 2;
				}
				else if (strnicmp( pchStart, "CHARACTER SET", pchEnd - pchStart) == 0 )
				{
					nKey = 3;
				}

				pchStart = pchPos + 1;
			}
			else if (*pchPos == ',' || *pchPos == ';' || *pchPos == '\0')
			{
				char chTemp = *pchPos;
				*pchPos = '\0';

				switch( nKey )
				{
				case 0:
					break;

				case 1:
					dpbcat(&m_pchDPB, &m_cbDPB, isc_dpb_sql_role_name, (void *)pchStart);
					break;

				case 2:
					g_dso_usDialect = atoi(pchStart);
					if( g_dso_usDialect == 3 )
					{
						HMODULE	hGds32;
						if (( (hGds32 = GetModuleHandle("FBCLIENT.DLL")) == NULL && (hGds32 = GetModuleHandle("GDS32.DLL")) == NULL) ||
							( g_dso_pIscEncodeSqlDate   = (ISCENCODESQLDATE)GetProcAddress(hGds32,   "isc_encode_sql_date") )  == NULL ||
							( g_dso_pIscEncodeSqlTime   = (ISCENCODESQLTIME)GetProcAddress(hGds32,   "isc_encode_sql_time") )  == NULL ||
							( g_dso_pIscEncodeTimestamp = (ISCENCODETIMESTAMP)GetProcAddress(hGds32, "isc_encode_timestamp") ) == NULL ||
							( g_dso_pIscDecodeSqlDate   = (ISCDECODESQLDATE)GetProcAddress(hGds32,   "isc_decode_sql_date") )  == NULL ||
							( g_dso_pIscDecodeSqlTime   = (ISCDECODESQLTIME)GetProcAddress(hGds32,   "isc_decode_sql_time") )  == NULL ||
							( g_dso_pIscDecodeTimestamp = (ISCDECODETIMESTAMP)GetProcAddress(hGds32, "isc_decode_timestamp") ) == NULL )
						{
							OutputDebugString( _T("SQL Dialect 3 specified, but related functions could not be loaded.  Reverting to SQL Dialect 1." ) );
							g_dso_usDialect = 1;
						}
					}
					break;

				case 3:
					dpbcat(&m_pchDPB, &m_cbDPB, isc_dpb_lc_ctype, (void *)pchStart);
				}

				if ((*pchPos = chTemp) == '\0')
					break;

				pchStart = pchPos + 1;
				nKey = 0;
			}
			pchPos++;
		}
	}

	if( SUCCEEDED( GetStringFromVariant( (char **)&pszValue, _MAX_PATH, NULL, GetIbPropertyValue( pIbProperties + AUTH_PASSWORD_INDEX ) ) ) )
	{
		dpbcat(&m_pchDPB, &m_cbDPB, isc_dpb_password, (void *)pszValue);
	}

	char *pchEnd = pszValue;

	if( SUCCEEDED( GetStringFromVariant( (char **)&pszValue, _MAX_PATH - 2, NULL, GetIbPropertyValue( pIbProperties + INIT_LOCATION_INDEX ) ) ) )
	{
		while (*pchEnd)
			pchEnd++;

		if( ( pszValue[0] == '\\' && pszValue[1] == '\\' ) ||
			( pszValue[0] == '/' && pszValue[1] == '/' ) )
		{
			//Handle Named-Pipes Syntax
			*pchEnd++ = '/';
		} 
		else if (pchEnd > pszValue && !(pchEnd[-1] == '@' || pchEnd[-1] == ':'))
		{
			*pchEnd++ = ':';
		}
	}

	if( SUCCEEDED( GetStringFromVariant( (char **)&pchEnd, _MAX_PATH, NULL, GetIbPropertyValue( pIbProperties + INIT_DATASOURCE_INDEX ) ) ) )
	{
		if (isc_attach_database(iscStatus, 0, pszValue, &m_iscDbHandle, m_cbDPB, m_pchDPB))
		{
			m_pInterBaseError->PostError(iscStatus);
			return DB_E_ERRORSOCCURRED;
		}
	}
	else
		return E_FAIL;
	
	m_dwStatus |= DSF_INITIALIZED;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CInterBaseDSO::Uninitialize(void)
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseDSO::Uninitialize()");

	if (m_cOpenSessions > 0)
		return DB_E_OBJECTOPEN;
	
	if (m_iscDbHandle)
	{
		isc_detach_database(NULL, &m_iscDbHandle);
	}

	m_dwStatus &= ~DSF_INITIALIZED;
	return S_OK;
}

HRESULT CInterBaseDSO::BeginMethod(
	REFIID	   riid)
{
	if (!(m_pInterBaseError ||
		 (m_pInterBaseError = new CInterBaseError)))
		return E_OUTOFMEMORY;

	m_pInterBaseError->ClearError(riid);

	return S_OK;
}

HRESULT WINAPI CInterBaseDSO::InternalQueryInterface(
	CInterBaseDSO *				pThis,
	const _ATL_INTMAP_ENTRY*	pEntries, 
	REFIID						iid, 
	void**						ppvObject)
{
	HRESULT    hr;
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseDSO::InternalQueryInterface()");

	DbgPrintGuid( iid );

	if( ppvObject == NULL )
	{
		hr = E_INVALIDARG;
	}
	else if( ( m_dwStatus & DSF_INITIALIZED ) == 0 &&
             ( IsEqualGUID(iid, IID_IDBCreateSession) ||
			   IsEqualGUID(iid, IID_IDBInfo ) ) )
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

HRESULT STDMETHODCALLTYPE CInterBaseDSO::GetProperties( 
        ULONG				   cPropertyIDSets,
        const DBPROPIDSET	   rgPropertyIDSets[],
        ULONG				  *pcPropertySets,
        DBPROPSET			 **prgPropertySets)
{
	CSmartLock	   sm(&m_cs);

	DbgFnTrace("CInterBaseDSO::GetProperties()");

	if (!cPropertyIDSets)
	{
		DBPROPIDSET	rgPropIDSets[] = 
				{{NULL, -1, EXPANDGUID(DBPROPSET_DATASOURCEINFO)},
				 {NULL, -1, EXPANDGUID(DBPROPSET_DBINIT)}};
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
        
HRESULT STDMETHODCALLTYPE CInterBaseDSO::GetPropertyInfo( 
        /* [in] */ ULONG cPropertyIDSets,
        /* [size_is][in] */ const DBPROPIDSET __RPC_FAR rgPropertyIDSets[],
        /* [out][in] */ ULONG __RPC_FAR *pcPropertyInfoSets,
        /* [size_is][size_is][out] */ DBPROPINFOSET __RPC_FAR *__RPC_FAR *prgPropertyInfoSets,
        /* [out] */ OLECHAR __RPC_FAR *__RPC_FAR *ppDescBuffer)
{
	CSmartLock	   sm(&m_cs);
	HRESULT		   hr = S_OK;
	IBPROPERTY	  *pIbProperties;
	IBPROPERTY	  *pIbPropertiesEnd;
	const DBPROPIDSET	  *pPropertyIDSet = rgPropertyIDSets;
	const DBPROPIDSET	  *pPropertyIDSetEnd = pPropertyIDSet + cPropertyIDSets;
	DBPROPIDSET	   rgSpecialPropertyIDSets[MAX_SPECIAL_PROPERTYSETS];
	ULONG		   cSpecialPropertyIDSets = 0;
	ULONG		   cDescriptions = 0;
	ULONG		   cProperties;
	ULONG		   cIbProperties;
	DBPROPINFO		*rgPropInfo;
	DBPROPINFO		*rgPropInfoEnd;
	DBPROPINFOSET	*rgPropInfoSet;
	bool		   fSpecial = false;
	CComObject<CInterBaseCommand> *pCommand = NULL;
	CComObject<CInterBaseSession> *pSession = NULL;

	DbgFnTrace("CInterBaseDSO::GetPropertyInfo()");

	// Initialize
	if (!g_pIMalloc && FAILED(CoGetMalloc(MEMCTX_TASK, &g_pIMalloc)))
		return E_FAIL;

	if( pcPropertyInfoSets )
		*pcPropertyInfoSets = 0;
	if( prgPropertyInfoSets )
		*prgPropertyInfoSets = NULL;
	if( ppDescBuffer )
		*ppDescBuffer = NULL;

	// Check Arguments
	if( ((cPropertyIDSets > 0) && !rgPropertyIDSets) ||
		!pcPropertyInfoSets || !prgPropertyInfoSets )
		return E_INVALIDARG;

	// If cPropertyIDSets == 0 then get the proper property sets
	if (cPropertyIDSets == 0)
	{
		DBPROPIDSET	  *rgPropIDSets = (DBPROPIDSET *)alloca(sizeof(DBPROPIDSET));
		
		cPropertyIDSets = 1;

		rgPropIDSets[0].cPropertyIDs = 0;
		rgPropIDSets[0].rgPropertyIDs = NULL;
		
		if (m_dwStatus & DSF_INITIALIZED)
			rgPropIDSets[0].guidPropertySet = DBPROPSET_DATASOURCEALL;
		else
			rgPropIDSets[0].guidPropertySet = DBPROPSET_DBINITALL;

		rgPropertyIDSets = rgPropIDSets;
	}
	
	//If any of the requested property sets are the special ALL ones then we need to expand then
	//to the build a list of the real property sets.
	//If there are more than MAX_SPECIAL_PROPERTYSETS then we will just ignore the extra ones since
	//it's an abuse case, but as we add more provider specific property sets we'll beed to increare
	//MAX_SPECIAL_PROPERTYSETS or do this a more elegant way.
	for ( ; pPropertyIDSet < pPropertyIDSetEnd; pPropertyIDSet++)
	{
		GUID const& guidSet = pPropertyIDSet->guidPropertySet;
		rgSpecialPropertyIDSets[cSpecialPropertyIDSets].rgPropertyIDs = NULL;
		rgSpecialPropertyIDSets[cSpecialPropertyIDSets].cPropertyIDs = 0;

		if (IsEqualGUID(guidSet, DBPROPSET_DATASOURCEALL))
		{
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_DATASOURCE;
		}
		else if (IsEqualGUID(guidSet, DBPROPSET_DATASOURCEINFOALL))
		{
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_DATASOURCEINFO;
		}
		else if (IsEqualGUID(guidSet, DBPROPSET_DBINITALL))
		{
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_DBINIT;
		}
		else if (IsEqualGUID(guidSet, DBPROPSET_SESSIONALL))
		{
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_SESSION;
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_IBASE_SESSION;
		}
		else if (IsEqualGUID(guidSet, DBPROPSET_ROWSETALL))
		{
			if (cSpecialPropertyIDSets < MAX_SPECIAL_PROPERTYSETS)
				rgSpecialPropertyIDSets[cSpecialPropertyIDSets++].guidPropertySet = DBPROPSET_ROWSET;
		}
		else if (cSpecialPropertyIDSets > 0)
		{
			//This property set is not special, but there are some special ones too.  This is an error.
			return E_INVALIDARG;
		}
	}
	if (fSpecial = (cSpecialPropertyIDSets > 0))
	{
		pPropertyIDSet = rgPropertyIDSets = rgSpecialPropertyIDSets;
		cPropertyIDSets = cSpecialPropertyIDSets;
		pPropertyIDSetEnd = pPropertyIDSet + cPropertyIDSets;
	}

	rgPropInfoSet = (DBPROPINFOSET *)g_pIMalloc->Alloc(cPropertyIDSets * sizeof(DBPROPINFOSET));
	if (!(*prgPropertyInfoSets = rgPropInfoSet))
		return E_OUTOFMEMORY;

	for (pPropertyIDSet = rgPropertyIDSets; pPropertyIDSet < pPropertyIDSetEnd; pPropertyIDSet++)
	{
		GUID const& guidPropSet = pPropertyIDSet->guidPropertySet;
		IBPROPCALLBACK pfnGetIbProperties = NULL;
		pIbProperties = NULL;
		cIbProperties = 0;
	
		rgPropInfoSet->guidPropertySet = GUID_NULL;
		rgPropInfoSet->cPropertyInfos = 0;
		rgPropInfoSet->rgPropertyInfos = NULL;

		if (IsEqualGUID(guidPropSet, DBPROPSET_DATASOURCE))
		{
			pfnGetIbProperties = _getIbProperties;
		}
		else if (IsEqualGUID(guidPropSet, DBPROPSET_DBINIT))
		{
			pfnGetIbProperties = _getIbProperties;
		}
		else if (IsEqualGUID(guidPropSet, DBPROPSET_DATASOURCEINFO))
		{
			pfnGetIbProperties = _getIbProperties;
		}
		else if (IsEqualGUID(guidPropSet, DBPROPSET_ROWSET))
		{
			if( NULL == ( pCommand = new CComObject<CInterBaseCommand> ) )
				return E_OUTOFMEMORY;

			pfnGetIbProperties = pCommand->_getIbProperties;
		}
		else if (IsEqualGUID(guidPropSet, DBPROPSET_SESSION))
		{
			if( NULL == ( pSession = new CComObject<CInterBaseSession> ) )
				return E_OUTOFMEMORY;

			pfnGetIbProperties = pSession->_getIbProperties;
		}
		else if (IsEqualGUID(guidPropSet, DBPROPSET_IBASE_SESSION))
		{
			if( NULL == ( pSession = new CComObject<CInterBaseSession> ) )
				return E_OUTOFMEMORY;

			pfnGetIbProperties = pSession->_getIbProperties;
		}

		if (pfnGetIbProperties != NULL)
		{
			pIbProperties = pfnGetIbProperties(guidPropSet, &cIbProperties);
			_ASSERTE(cIbProperties > 0);
		}

		cProperties = fSpecial ? cIbProperties : pPropertyIDSet->cPropertyIDs;

		rgPropInfo = (DBPROPINFO *)g_pIMalloc->Alloc(cProperties * sizeof(DBPROPINFO));
		if (!(rgPropInfoSet->rgPropertyInfos = rgPropInfo))
		{
			hr = E_OUTOFMEMORY;
			goto cleanup;
		}
		rgPropInfoSet->guidPropertySet = guidPropSet;
		rgPropInfoSet->cPropertyInfos = cProperties;
		pIbPropertiesEnd = pIbProperties + cIbProperties;

		if (fSpecial)
		{
			for ( ; pIbProperties < pIbPropertiesEnd; pIbProperties++, rgPropInfo++)
			{
				rgPropInfo->dwFlags = pIbProperties->dwFlags;
				rgPropInfo->dwPropertyID = pIbProperties->dwID;
				rgPropInfo->pwszDescription = pIbProperties->wszName;
				rgPropInfo->vtType = pIbProperties->vt;
				VariantInit(&rgPropInfo->vValues);
			}
		}
		else
		{
			ULONG		  *pPropID = pPropertyIDSet->rgPropertyIDs;
			ULONG		  *pPropIDEnd = pPropID + pPropertyIDSet->cPropertyIDs;

			for ( ; pPropID < pPropIDEnd; pPropID++)
			{
				VariantInit(&rgPropInfo->vValues);
				rgPropInfo->dwPropertyID = *pPropID;
				for ( ; pIbProperties < pIbPropertiesEnd; pIbProperties++)
				{
					if (pIbProperties->dwID == *pPropID)
					{
						rgPropInfo->dwFlags = pIbProperties->dwFlags;
						rgPropInfo->pwszDescription = pIbProperties->wszName;
						rgPropInfo->vtType = pIbProperties->vt;
						break;
					}
				}
				if (pIbProperties == pIbPropertiesEnd)
				{
					rgPropInfo->dwFlags = DBPROPFLAGS_NOTSUPPORTED;
					rgPropInfo->pwszDescription = NULL;
					rgPropInfo->vtType = VT_EMPTY;
				}
				rgPropInfo++;
			}
		}
		cDescriptions += cProperties;
		rgPropInfoSet++;
	}

	//Now that we have all the properties, we need to copy the descriptions from
	//the private memory into a buffer we allocate for the consumer.
	if ((*pcPropertyInfoSets = (rgPropInfoSet - *prgPropertyInfoSets)) && ppDescBuffer)
	{
		ULONG	   i;
		LPOLESTR   pName;

		if (!(pName = (LPOLESTR)g_pIMalloc->Alloc(cDescriptions * MAX_PROPERTY_NAME_LENGTH * sizeof(OLECHAR))))
		{
			hr = E_OUTOFMEMORY;
			goto cleanup;
		}
		*ppDescBuffer = pName;

		for (i = 0; i < *pcPropertyInfoSets; i++)
		{
			rgPropInfo = (*prgPropertyInfoSets)[i].rgPropertyInfos;
			rgPropInfoEnd = rgPropInfo + (*prgPropertyInfoSets)[i].cPropertyInfos;
			for ( ; rgPropInfo < rgPropInfoEnd; rgPropInfo++)
			{
				if( rgPropInfo->pwszDescription != NULL )
					wcscpy(pName, rgPropInfo->pwszDescription);
				else
					*pName = L'\0';

				rgPropInfo->pwszDescription = pName;
				_ASSERTE(wcslen(pName) < MAX_PROPERTY_NAME_LENGTH);
				pName += MAX_PROPERTY_NAME_LENGTH;
			}
		}
	}

cleanup:
	if( pCommand )
		delete pCommand;

	if( pSession )
		delete pSession;

	return hr;
}
	    
HRESULT STDMETHODCALLTYPE CInterBaseDSO::SetProperties( 
        ULONG		   cPropertySets,
        DBPROPSET	   rgPropertySets[])
{
	CSmartLock sm(&m_cs);

	DbgFnTrace("CInterBaseDSO::SetProperties()");

	return SetIbProperties(cPropertySets, rgPropertySets, _getIbProperties);
}

//
// ISupportErrorInfo methods
//
STDMETHODIMP CInterBaseDSO::InterfaceSupportsErrorInfo
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

BOOL CALLBACK DialogProc(
		HWND	   hWnd, 
		UINT	   uMsg, 
		WPARAM	   wParam, 
		LPARAM	   lParam)
{
	USES_CONVERSION;
	static	bool fIsRemote = false;
	static IBPROPERTY *pIbProperties = NULL;

	switch(uMsg)
	{
	case WM_INITDIALOG:
		{	
			pIbProperties = (IBPROPERTY *)lParam;
			SetWindowText( GetDlgItem( hWnd, IDC_USERNAME ), (char *)OLE2A(V_BSTR( &((pIbProperties + AUTH_USERID_INDEX)->vtVal) ) ) );
			SetWindowText( GetDlgItem( hWnd, IDC_PASSWORD ), (char *)OLE2A(V_BSTR( &((pIbProperties + AUTH_PASSWORD_INDEX)->vtVal) ) ) );
			SetWindowText( GetDlgItem( hWnd, IDC_DATABASE ), (char *)OLE2A(V_BSTR( &((pIbProperties + INIT_DATASOURCE_INDEX)->vtVal) ) ) );
			SetWindowText( GetDlgItem( hWnd, IDC_OPTIONS ), (char *)OLE2A(V_BSTR( &((pIbProperties + INIT_PROVIDERSTRING_INDEX)->vtVal) ) ) );

			fIsRemote = (V_BSTR( &((pIbProperties+ INIT_LOCATION_INDEX)->vtVal ))[0] != L'' );

			SetWindowText( GetDlgItem( hWnd, IDC_SERVER ), (char *)OLE2A(V_BSTR( &((pIbProperties + INIT_LOCATION_INDEX)->vtVal) ) ) );

			SendMessage(GetDlgItem(hWnd, IDC_LOCAL), BM_SETCHECK, !fIsRemote, 0);
			SendMessage(GetDlgItem(hWnd, IDC_REMOTE), BM_SETCHECK, fIsRemote, 0);
			EnableWindow(GetDlgItem(hWnd, IDC_SERVER), fIsRemote);
			EnableWindow(GetDlgItem(hWnd, IDC_PROTOCOL), fIsRemote);
		}
		return 1;

	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDC_LOCAL:
		case IDC_REMOTE:
			{
				fIsRemote = ( LOWORD(wParam) == IDC_REMOTE );
				SendMessage(GetDlgItem(hWnd, IDC_LOCAL), BM_SETCHECK, !fIsRemote, 0);
				SendMessage(GetDlgItem(hWnd, IDC_REMOTE), BM_SETCHECK, fIsRemote, 0);
				EnableWindow(GetDlgItem(hWnd, IDC_SERVER), fIsRemote);
				EnableWindow(GetDlgItem(hWnd, IDC_PROTOCOL), fIsRemote);
			}
			return 1;

		case IDOK:
			{
				char szValue[512];

				GetWindowText( GetDlgItem(hWnd, IDC_USERNAME), szValue, 32 );
				SetIbPropertyValue( (pIbProperties + AUTH_USERID_INDEX), & CComVariant( A2W( szValue ) ) );

				GetWindowText( GetDlgItem(hWnd, IDC_PASSWORD), szValue, 32 );
				SetIbPropertyValue( (pIbProperties + AUTH_PASSWORD_INDEX), & CComVariant( A2W( szValue ) ) );

				GetWindowText( GetDlgItem(hWnd, IDC_DATABASE), szValue, 512 );
				SetIbPropertyValue( (pIbProperties+ INIT_DATASOURCE_INDEX), & CComVariant( A2W( szValue ) ) );

				GetWindowText( GetDlgItem(hWnd, IDC_OPTIONS), szValue, 512 );
				SetIbPropertyValue( (pIbProperties + INIT_PROVIDERSTRING_INDEX), & CComVariant( A2W( szValue ) ) );

				if( fIsRemote )
				{
					GetWindowText( GetDlgItem(hWnd, IDC_SERVER), szValue, 32 );
				}
				else
				{
					szValue[0] = '\0';
				}
				SetIbPropertyValue( (pIbProperties + INIT_LOCATION_INDEX), & CComVariant( A2W( szValue ) ) );

				EndDialog((HWND)hWnd, TRUE);
			}
			return 1;

		case IDCANCEL:
			EndDialog((HWND)hWnd, FALSE);
			return 1;
		}
	}
	return 0;
}