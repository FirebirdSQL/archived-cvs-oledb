/*
 * Firebird Open Source OLE DB driver
 *
 * PROPERTY.CPP - Source file for various OLE DB property related functions.
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
#define DEFINE_IBOLEDB_PROPERTIES
#include "../inc/property.h"

VARIANTARG *GetIbPropertyValue(IBPROPERTY *pIbProperty)
{
	VARIANTARG *v = &(pIbProperty->vtVal);

	if (!(pIbProperty->fIsInitialized || pIbProperty->fIsSet)) 
	{	
		VariantInit(v);
		pIbProperty->fIsInitialized = true;
		v->vt = pIbProperty->vt;

		switch (pIbProperty->vt) 
		{
			case VT_BOOL:
				v->boolVal = pIbProperty->dwVal ? VARIANT_TRUE : VARIANT_FALSE; 
				break;
			case VT_BSTR: 
				v->bstrVal = SysAllocString(pIbProperty->wszVal); 
				break;
			default: 
				v->lVal = pIbProperty->dwVal;
		}
	}
	return v;
}

DBPROPSTATUS SetIbPropertyValue(IBPROPERTY *pIbProperty, VARIANT *pvNewValue)
{
	VARIANT *pvOldValue;

	//
	// If the new value is a VT_EMPTY then return the property to its default state
	//
	if (pvNewValue->vt == VT_EMPTY)
	{
		if (pIbProperty->fIsSet || pIbProperty->fIsInitialized)
		{
			VariantClear(&pIbProperty->vtVal);
			pIbProperty->fIsInitialized = false;
			pIbProperty->fIsSet = false;
		}

		return DBPROPSTATUS_OK;
	}

	//
	// If the type of the new value is not right, then error
	//
	if( V_VT( pIbProperty ) != V_VT( pvNewValue ) )
		return DBPROPSTATUS_BADVALUE;

	//
	// If the property is writeable then set its value
	//
	if (pIbProperty->dwFlags & DBPROPFLAGS_WRITE)
	{
		if (pIbProperty->fIsSet)
			VariantClear(&pIbProperty->vtVal);
		else
			VariantInit(&pIbProperty->vtVal);
		
		VariantCopy( &pIbProperty->vtVal, pvNewValue );
		pIbProperty->fIsSet = TRUE;

		return DBPROPSTATUS_OK;
	}

	//
	// The property is read-only.  We can set the value to its present value, but not a new value
	//
	pvOldValue = GetIbPropertyValue(pIbProperty);

	switch(pIbProperty->vt)
	{
		case VT_BOOL:
			if( V_BOOL( pvOldValue ) == V_BOOL( pvNewValue ) )
				return DBPROPSTATUS_OK;
			
			return DBPROPSTATUS_NOTSETTABLE;

		case VT_I4:
			if( V_I4( pvOldValue ) == V_I4( pvNewValue ) )
				return DBPROPSTATUS_OK;

			return DBPROPSTATUS_NOTSETTABLE;

		case VT_BSTR:
			if( ! wcscmp( V_BSTR( pvOldValue ), V_BSTR( pvNewValue ) ) )
				return DBPROPSTATUS_OK;

			return DBPROPSTATUS_NOTSETTABLE;
	}

	return DBPROPSTATUS_BADVALUE;
}

IBPROPERTY *LookupIbProperty(
		IBPROPERTY	  *pIbProperties, 
		DWORD		   cIbProperties, 
		DWORD		   dwPropID)
{
	IBPROPERTY	  *pIbProperty = pIbProperties;
	IBPROPERTY	  *pIbPropertyEnd = pIbProperty + cIbProperties;

	for ( ; pIbProperty < pIbPropertyEnd; pIbProperty++)
	{
		if (pIbProperty->dwID == dwPropID)
			return pIbProperty;
	}

	return NULL;
}

HRESULT GetIbProperties(
	const ULONG		   cPropertyIDSets,
	const DBPROPIDSET  rgPropertyIDSets[],
	ULONG			  *pcPropertySets,
	DBPROPSET		 **prgPropertySets,
	IBPROPCALLBACK	   pfnGetIbProperties)
{
	HRESULT		   hr = S_OK;
	bool		   fGetAllProperties = false;
	DBPROPSET	  *pPropertySets = NULL;
	ULONG		   cPropertySets = 0;
	const DBPROPIDSET	  *pPropertyIDSet = rgPropertyIDSets;
	const DBPROPIDSET	  *pPropertyIDSetEnd = pPropertyIDSet + cPropertyIDSets;
	ULONG		   j;
	IBPROPERTY	  *pIbProperties;
	ULONG		   cIbProperties;

	if (!(pPropertySets = (DBPROPSET *)g_pIMalloc->Alloc(cPropertyIDSets * sizeof(DBPROPSET))))
	{
		hr = E_OUTOFMEMORY;
		goto cleanup;
	}

	for ( ; pPropertyIDSet < pPropertyIDSetEnd; pPropertyIDSet++)
	{
		if (pIbProperties = pfnGetIbProperties(pPropertyIDSet->guidPropertySet, &cIbProperties))
		{
			if (pPropertyIDSet->cPropertyIDs == -1 &&
				pPropertyIDSet->rgPropertyIDs == NULL)
			{
				fGetAllProperties = TRUE;
				pPropertySets[cPropertySets].cProperties = cIbProperties;
			}
			else
			{
				pPropertySets[cPropertySets].cProperties = pPropertyIDSet->cPropertyIDs;
			}
			pPropertySets[cPropertySets].rgProperties = (DBPROP *)g_pIMalloc->Alloc(pPropertySets[cPropertySets].cProperties * sizeof(DBPROP));
			pPropertySets[cPropertySets].guidPropertySet = pPropertyIDSet->guidPropertySet;

			if (fGetAllProperties)
			{
				for (j = 0; j < cIbProperties; j++)
				{
					char szNumber[11];
					itoa( pIbProperties[j].dwID, szNumber, 10 );
					pPropertySets[cPropertySets].rgProperties[j].colid = DB_NULLID;
					pPropertySets[cPropertySets].rgProperties[j].dwOptions = 0;
					pPropertySets[cPropertySets].rgProperties[j].dwPropertyID = pIbProperties[j].dwID;
					pPropertySets[cPropertySets].rgProperties[j].dwStatus = DBSTATUS_S_OK;
					VariantInit(&pPropertySets[cPropertySets].rgProperties[j].vValue);
					VariantCopy(&pPropertySets[cPropertySets].rgProperties[j].vValue, GetIbPropertyValue(&pIbProperties[j]));
				}
			}
			else
			{
				IBPROPERTY	  *pIbProperty;
				DBPROPSTATUS   dwStatus;

				for (j = 0; j < pPropertyIDSet->cPropertyIDs; j++)
				{
					pPropertySets[cPropertySets].rgProperties[j].colid = DB_NULLID;
					pPropertySets[cPropertySets].rgProperties[j].dwPropertyID = pPropertyIDSet->rgPropertyIDs[j];
					VariantInit(&pPropertySets[cPropertySets].rgProperties[j].vValue);

					if (pIbProperty = LookupIbProperty(pIbProperties, cIbProperties, pPropertyIDSet->rgPropertyIDs[j]))
					{
						dwStatus = DBPROPSTATUS_OK;	
						pPropertySets[cPropertySets].rgProperties[j].dwOptions = 0;
						VariantCopy(&pPropertySets[cPropertySets].rgProperties[j].vValue, GetIbPropertyValue(pIbProperty));
					}
					else
					{
						dwStatus = DBPROPSTATUS_NOTSUPPORTED;
					}
					pPropertySets[cPropertySets].rgProperties[j].dwStatus = dwStatus;
				}
			}
			cPropertySets++;
		}
		else if( DBPROPSET_PROPERTIESINERROR == pPropertyIDSet->guidPropertySet )
		{
			if( pPropertySets )
				g_pIMalloc->Free(pPropertySets);

			*pcPropertySets = 0;
			*prgPropertySets = NULL;

			if( cPropertyIDSets == 1 )
				return S_OK;

			return DB_E_ERRORSOCCURRED;
		}
	}

cleanup:
	if (FAILED(hr) || cPropertySets == 0)
	{
		if (!cPropertySets)
			hr = DB_E_ERRORSOCCURRED;
		if (pPropertySets)
			g_pIMalloc->Free(pPropertySets);
		cPropertySets = 0;
		pPropertySets = NULL;
	}
	else if (cPropertySets < cPropertyIDSets)
	{
		hr = DB_S_ERRORSOCCURRED;
	}

	*pcPropertySets = cPropertySets;
	*prgPropertySets = pPropertySets;

	return hr;
}

HRESULT SetIbProperties( 
	ULONG		   cPropertySets,
	DBPROPSET	   rgPropertySets[],
	IBPROPCALLBACK pfnGetIbProperties)
{
	HRESULT		   hr = S_OK;
	ULONG		   cPropertiesSet = 0;
	ULONG		   cProperties = 0;
	DBPROPSET	  *DbPropSet = rgPropertySets;
	DBPROPSET	  *DbPropSetEnd = DbPropSet + cPropertySets;
	IBPROPERTY	  *pIbProperty;
	IBPROPERTY	  *pIbProperties;
	ULONG		   cIbProperties;

	for ( ; DbPropSet < DbPropSetEnd; DbPropSet++)
	{
		DBPROP *DbProp = DbPropSet->rgProperties;
		DBPROP *DbPropEnd = DbProp + DbPropSet->cProperties;
		
		cProperties += DbPropSet->cProperties;
		
		pIbProperties = pfnGetIbProperties(DbPropSet->guidPropertySet, &cIbProperties);

		for ( ; DbProp < DbPropEnd; DbProp++)
		{
			if ((pIbProperty = LookupIbProperty(pIbProperties, cIbProperties, DbProp->dwPropertyID)))
			{
				DbProp->dwStatus = SetIbPropertyValue( pIbProperty, &( DbProp->vValue ) );
			}
			else
			{
				DbProp->dwStatus = DBPROPSTATUS_NOTSUPPORTED;
				hr = DB_S_ERRORSOCCURRED;
			}
			if (DbProp->dwStatus == DBPROPSTATUS_OK)
				cPropertiesSet++;

		}
	}
	if( cProperties && !cPropertiesSet )
	{
		hr = DB_E_ERRORSOCCURRED;
	}

	return hr;
}

HRESULT ClearIbProperties (
		IBPROPERTY	  *pIbProperty, 
		DWORD		   cIbProperties)
{
	IBPROPERTY	  *pIbPropertyEnd = pIbProperty + cIbProperties;

	for ( ; pIbProperty < pIbPropertyEnd; pIbProperty++)
	{
		if (pIbProperty->fIsInitialized || pIbProperty->fIsSet)
		{
			VariantClear(&(pIbProperty->vtVal));
			pIbProperty->fIsInitialized = false;
			pIbProperty->fIsSet = false;
		}
	}

	return S_OK;
}
