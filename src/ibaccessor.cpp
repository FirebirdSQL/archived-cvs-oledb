/*
 * Firebird Open Source OLE DB driver
 *
 * IBACCESSOR.CPP - Include file for internal accessor methods.
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
#include "../inc/convert.h"
#include "../inc/ibaccessor.h"
#include <objbase.h>

HRESULT CreateIbAccessor(
		DBACCESSORFLAGS	   dwAccessorFlags,
		ULONG			   cBindings,
		const DBBINDING	   rgBindings[],
		ULONG			   cbRowSize,
		PIBACCESSOR		  *ppIbAccessor,
		DBBINDSTATUS	   rgStatus[],
		ULONG			   cColumns,
		IBCOLUMNMAP		  *pIbColumnMap)
{
	const DBBINDING	  *pBinding = rgBindings;
	const DBBINDING	  *pBindingEnd = pBinding + cBindings;
	PIBACCESSOR		   pNewIbAccessor;
	DBSTATUS		   wStatus;
	DWORD			   dwParamIO;
	ULONG			   nErr = 0;

	if (!ppIbAccessor)
		return E_INVALIDARG;

	*ppIbAccessor = NULL;

	if (!cBindings)
		return S_OK;

	for( ; pBinding < pBindingEnd; pBinding++ )
	{
		wStatus = DBBINDSTATUS_OK;

		if( ( pBinding->iOrdinal < 1 ) ||
			( cColumns && pBinding->iOrdinal > cColumns ) )
		{
			nErr++;
			wStatus = DBBINDSTATUS_BADORDINAL;
		}
		else if( ! ( pBinding->dwPart && ( DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS ) ) )
		{
			nErr++;
			wStatus = DBBINDSTATUS_BADBINDINFO;
		}
		else if( pBinding->wType == DBTYPE_IUNKNOWN && (
				 ! ( pIbColumnMap && (pIbColumnMap[pBinding->iOrdinal - 1].dwFlags & DBCOLUMNFLAGS_ISLONG )) ||
				 ( pBinding->pObject != NULL &&
			       pBinding->pObject->iid != IID_ISequentialStream ) ) )
		{
			nErr++;
			wStatus = DBBINDSTATUS_NOINTERFACE;
		}

		dwParamIO |= pBinding->eParamIO;

		if( rgStatus != NULL )
			rgStatus[pBinding - rgBindings] = wStatus;
	}

	if (nErr == cBindings)
	{
		return DB_E_ERRORSOCCURRED;
	}

	pNewIbAccessor = new IBACCESSOR[IBACCESSOR_LENGTH(cBindings)];
	if( pNewIbAccessor == NULL )
		return E_OUTOFMEMORY;

	pNewIbAccessor->dwFlags = dwAccessorFlags;
	pNewIbAccessor->cBindInfo = cBindings;
	pNewIbAccessor->ulRefCount = 1;
	pNewIbAccessor->pNext = NULL;
	pNewIbAccessor->dwParamIO = dwParamIO;

	memcpy( pNewIbAccessor->rgBindings, rgBindings, cBindings * sizeof(DBBINDING) );

	*ppIbAccessor = pNewIbAccessor;

	return S_OK;
}

HRESULT ReleaseIbAccessor(
		PIBACCESSOR	   pIbAccessor, 
		ULONG		  *pcRefCount,
		PIBACCESSOR	  *ppIbAccessorList)
{
	ULONG		   ulRefCount;

	if( ( ulRefCount = --pIbAccessor->ulRefCount ) == 0 )
	{
		if( pIbAccessor == *ppIbAccessorList )
		{
			*ppIbAccessorList = pIbAccessor->pNext;
			delete[] pIbAccessor;
		}
		else
		{
			PIBACCESSOR	pThis;
			PIBACCESSOR	pLast;

			for( pLast = *ppIbAccessorList, pThis = pLast->pNext; 
				 pThis != NULL; 
				 pLast = pThis, pThis = pThis->pNext )
			{
				if( pThis == pIbAccessor )
				{
					pLast->pNext = pThis->pNext;
					delete[] pThis;
					break;
				}
			}
		}
	}

	*pcRefCount = ulRefCount;

	return S_OK;
}

HRESULT GetIbBindings(
		PIBACCESSOR		   pIbAccessor,
		DBACCESSORFLAGS	  *pdwAccessorFlags,
		ULONG			  *pcBindings,
		DBBINDING		 **prgBindings)
{
	DBBINDING	  *pBindings;
	ULONG		   cBindings;
	ULONG		   cbBindings;

	if( pdwAccessorFlags == NULL || pcBindings == NULL || prgBindings == NULL )
		return E_INVALIDARG;

	cBindings = pIbAccessor->cBindInfo;
	cbBindings = cBindings * sizeof(DBBINDING);

	if( (pBindings = (DBBINDING *)g_pIMalloc->Alloc( cbBindings ) ) == NULL )
		return E_OUTOFMEMORY;


	memcpy( pBindings, pIbAccessor->rgBindings, cbBindings );

	*pdwAccessorFlags = pIbAccessor->dwFlags;
	*prgBindings = pBindings;
	*pcBindings = cBindings;

	return S_OK;
}

HRESULT GetIbColumnInfo(
		ULONG		  *pcColumns, 
		DBCOLUMNINFO **prgInfo, 
		OLECHAR		 **ppStringsBuffer,
		ULONG		   cColumns,
		XSQLDA		  *pSqlda,
		IBCOLUMNMAP	  *pIbColumnMap)
{
	HRESULT		   hr = S_OK;
	ULONG		   iOrdinal;
	DBCOLUMNINFO  *pColumnInfo;
	DBCOLUMNINFO  *pColumnInfoBuffer;
	OLECHAR		  *pwszStringBuffer;
	OLECHAR		  *pwszName;

	if (!(pcColumns && prgInfo && ppStringsBuffer)) 
		return E_INVALIDARG;

	_ASSERT(pSqlda);
	_ASSERT(pIbColumnMap);

	pColumnInfoBuffer = (DBCOLUMNINFO *)g_pIMalloc->Alloc(cColumns * sizeof(DBCOLUMNINFO));
	if (!(pColumnInfo = pColumnInfoBuffer))
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	pwszName = (OLECHAR *)g_pIMalloc->Alloc(cColumns * MAX_SQL_NAME_SIZE * 2 * sizeof(WCHAR));
	if (!(pwszStringBuffer = pwszName))
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	for (iOrdinal = 0; iOrdinal < cColumns; iOrdinal++)
	{
		//Set the easy stuff first
		pColumnInfo->pTypeInfo = NULL;
		pColumnInfo->iOrdinal = iOrdinal + 1;
		pColumnInfo->dwFlags = pIbColumnMap[iOrdinal].dwFlags;
		pColumnInfo->ulColumnSize = pIbColumnMap[iOrdinal].ulColumnSize;
		pColumnInfo->bPrecision = pIbColumnMap[iOrdinal].bPrecision;
		pColumnInfo->bScale = pIbColumnMap[iOrdinal].bScale;
		pColumnInfo->wType = pIbColumnMap[iOrdinal].wType;

		//Write the alias name into the buffer
		if (pIbColumnMap[iOrdinal].wszAlias[0])
		{
			pColumnInfo->pwszName = pwszName;

			wcscpy(pwszName, pIbColumnMap[iOrdinal].wszAlias);

			pwszName += MAX_SQL_NAME_SIZE;
		}
		else
			pColumnInfo->pwszName = NULL;

		//Initialize the DBID
		pColumnInfo->columnid = pIbColumnMap[ iOrdinal ].columnid;
		if( pColumnInfo->columnid.eKind == DBKIND_GUID_NAME ||
			pColumnInfo->columnid.eKind == DBKIND_NAME )
		{
			pColumnInfo->columnid.uName.pwszName = pwszName;

			wcscpy(pwszName, pIbColumnMap[iOrdinal].wszName);
			
			pwszName += ( MAX_SQL_NAME_SIZE );
		}


		if (pSqlda->sqlvar[iOrdinal].sqlname_length != 0)
			pColumnInfo->dwFlags |= DBCOLUMNFLAGS_WRITEUNKNOWN;
		if( NULLABLE( pSqlda->sqlvar[iOrdinal].sqltype ) )
			pColumnInfo->dwFlags |= ( DBCOLUMNFLAGS_ISNULLABLE | DBCOLUMNFLAGS_MAYBENULL );

		pColumnInfo++;
	}

cleanup:
	*pcColumns = pColumnInfo - pColumnInfoBuffer;

	if (pColumnInfoBuffer == pColumnInfo)
		g_pIMalloc->Free(pColumnInfoBuffer);
	else
		*prgInfo = pColumnInfoBuffer;

	if (pwszStringBuffer == pwszName)
		g_pIMalloc->Free(pwszStringBuffer);
	else
		*ppStringsBuffer = pwszStringBuffer;

	return hr;
}

HRESULT MapIbColumnIDs(
		ULONG		   cColumnIDs, 
		const DBID	   rgColumnIDs[], 
		ULONG		   rgColumns[],
		ULONG		   cColumns,
		IBCOLUMNMAP	  *pIbColumnMap)
{
	ULONG		   i;
	ULONG		   j;
	DBKIND      eKind;
	DBKIND      eIbKind;
	LPOLESTR	   pwszName;
	LPOLESTR	   pwszIbName;

	if (!rgColumns || (cColumnIDs && !rgColumnIDs))
		return E_INVALIDARG;

	_ASSERT(pIbColumnMap);

	for( i = 0; i < cColumnIDs; i++ )
	{
		rgColumns[ i ] = DB_INVALIDCOLUMN;
		eKind = rgColumnIDs[ i ].eKind;
		pwszName = rgColumnIDs[i].uName.pwszName;
		for( j = 0; j < cColumns; j++ )
		{
			eIbKind = pIbColumnMap[ j ].columnid.eKind;
			pwszIbName = pIbColumnMap[j].columnid.uName.pwszName;

			if( ( eKind == DBKIND_GUID_NAME ||
				  eKind == DBKIND_NAME ||
				  eKind == DBKIND_PGUID_NAME ) &&
				( eIbKind == DBKIND_GUID_NAME ||
				  eIbKind == DBKIND_NAME ||
				  eIbKind == DBKIND_PGUID_NAME ) )
			{
				if( wcscmp( pwszName, pwszIbName ) == 0 )
				{
					rgColumns[i] = pIbColumnMap[j].iOrdinal;
					break;
				}
			}
			else if( ( eKind == DBKIND_GUID_PROPID ||
				       eKind == DBKIND_PGUID_PROPID ||
					   eKind == DBKIND_PROPID ) &&
					 ( eIbKind == DBKIND_GUID_PROPID ||
					   eIbKind == DBKIND_PGUID_PROPID ||
					   eIbKind == DBKIND_PROPID ) )
			{
				if( pwszName == pwszIbName )
				{
					rgColumns[i] = pIbColumnMap[j].iOrdinal;
					break;
				}
			}
			else if( IsEqualGUID( rgColumnIDs[ i ].uGuid.guid, pIbColumnMap[ j ].columnid.uGuid.guid ) )
			{
				rgColumns[i] = pIbColumnMap[j].iOrdinal;
				break;
			}
		}
	}
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
// SetSqlDaPointers
///////////////////////////////////////////////////////////////////////////////
// Fix-up the sqlda members, sqldata and sqlind to point to the proper offset
// in the current data buffer.
///////////////////////////////////////////////////////////////////////////////
HRESULT SetSqldaPointers(
		XSQLDA		  *pSqlda,
		IBCOLUMNMAP	  *pIbColumnMap,
		void		  *pvData )
{
	ULONG	   i;
	ULONG	   cColumns = pSqlda->sqld;

	for (i = 0; i < cColumns; i++)
	{
		pSqlda->sqlvar[i].sqldata =  (char *)( (BYTE *)pvData + pIbColumnMap[i].obData );
		pSqlda->sqlvar[i].sqlind  = (short *)( (BYTE *)pvData + pIbColumnMap[i].obStatus );
	}

	return S_OK;
}