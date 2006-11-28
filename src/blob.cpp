/*
 * Firebird Open Source OLE DB driver
 *
 * BLOB.CPP - Source file for BLOB handling methods.
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
#include "../inc/blob.h"

CInterBaseBlobStream::CInterBaseBlobStream() :
	m_iscBlobHandle(NULL),
	m_fNullTerminate(false)
{
}

CInterBaseBlobStream::~CInterBaseBlobStream()
{
	if(m_iscBlobHandle)
		isc_close_blob(m_iscStatus, &m_iscBlobHandle);
}

CInterBaseBlobStream::Initialize(
		isc_blob_handle	   iscBlobHandle,
		bool			   fNullTerminate)
{
	m_iscBlobHandle = iscBlobHandle;
	m_fNullTerminate = fNullTerminate;
}

//
// ISequentialStream methods
//
HRESULT STDMETHODCALLTYPE CInterBaseBlobStream::Read(
		void	  *pvBuffer, 
		ULONG	   cbBuffer, 
		ULONG	  *pcbBytesRead)
{
	DBSTATUS   sc = DBSTATUS_S_OK;
	char	  *pchBuffer = (char *)pvBuffer;
	char	  *pchBufferEnd = pchBuffer + cbBuffer;
	USHORT	   usBytesRead;

	if (pvBuffer == NULL)
		return STG_E_INVALIDPOINTER;

	while (pchBuffer < pchBufferEnd &&
		   (!isc_get_segment(
			m_iscStatus,
			&m_iscBlobHandle,
			&usBytesRead,
			MIN(pchBufferEnd - pchBuffer, SHRT_MAX),
			pchBuffer) || 
			m_iscStatus[1] == isc_segment))
	{
		pchBuffer += usBytesRead;
	}

	/*
	if (m_fNullTerminate && pchBuffer < pchBufferEnd)
		*pchBuffer++ = '\0';
	*/

	if (pcbBytesRead)
		*pcbBytesRead = pchBuffer - (char *)pvBuffer;

	return sc;
}

HRESULT STDMETHODCALLTYPE CInterBaseBlobStream::Write(
		void const	  *pvBuffer, 
		ULONG		   cbBuffer, 
		ULONG		  *pcbBytesWritten)
{
	if (pvBuffer == NULL)
		return STG_E_INVALIDPOINTER;

	return S_FALSE;
}

DBSTATUS GetArray(
		isc_db_handle	  *piscDbHandle,
		isc_tr_handle	  *piscTransHandle,
		XSQLVAR			  *pSqlVar,
		BYTE			  *sqldata,
		short			  *sqlind,
		DBTYPE			   wType,
		BYTE			  *pbValue,
		DWORD			  *pdwLen,
		DWORD			  *pdwFlags,
		void			  **ppvdesc)
{
	UINT			   i;
	HRESULT			   hr;
	ISC_STATUS		   iscStatus[20];
	ISC_ARRAY_DESC	  *pdesc = *(ISC_ARRAY_DESC **)ppvdesc;
	VARTYPE			   vt;
	UINT			   cDims;
	SAFEARRAY		  *psa = NULL;
	SAFEARRAYBOUND	   rgsabound[16];
	void			  *pvData;
	long			   nSize = 0;

	if( wType != DBTYPE_VARIANT )
		return DBSTATUS_E_CANTCONVERTVALUE;

	if( NULLABLE( pSqlVar->sqltype ) && *sqlind )
	{
		if (*pdwFlags & DBPART_VALUE)
		{
			VariantInit((VARIANTARG *)pbValue);
			((VARIANTARG *)pbValue)->vt = VT_NULL;
		}
		return DBSTATUS_S_ISNULL;
	}
	else if (!(*pdwFlags & DBPART_VALUE))
		return DBSTATUS_S_OK;
	else
		VariantInit((VARIANTARG *)pbValue);

	if (!pdesc)
	{
		pdesc = new ISC_ARRAY_DESC[1];

		if (isc_array_lookup_bounds(
				iscStatus,
				piscDbHandle,
				piscTransHandle,
				pSqlVar->relname,
				pSqlVar->sqlname,
				pdesc))
		{
			delete[] pdesc;
			return DBSTATUS_E_CANTCONVERTVALUE;
		}
		*ppvdesc = (void *)pdesc;
	}

	cDims = pdesc->array_desc_dimensions;

	for (i = 0; i < cDims; i++)
	{
		rgsabound[i].lLbound = pdesc->array_desc_bounds[i].array_bound_lower;
		rgsabound[i].cElements = pdesc->array_desc_bounds[i].array_bound_upper - rgsabound[i].lLbound + 1;
		nSize += rgsabound[i].cElements;
	}
	nSize *= pdesc->array_desc_length;

	switch (pdesc->array_desc_dtype)
	{
	case blr_short:
		vt = VT_I2;
		break;
	case blr_long:
		vt = VT_I4;
		break;
	case blr_float:
		vt = VT_R4;
		break;
	case blr_double:
		vt = VT_R8;
		break;
	case blr_date:
		vt = VT_DATE;
		break;
	case SQL_TEXT:
	case SQL_VARYING:
		vt = VT_BSTR;
		break;
	default:
		return DBSTATUS_E_CANTCONVERTVALUE;
	}

	psa = SafeArrayCreate(vt, cDims, rgsabound);
	hr = SafeArrayAccessData(psa, &pvData);

	if (isc_array_get_slice(
			iscStatus, 
			piscDbHandle, 
			piscTransHandle, 
			(ISC_QUAD *)sqldata, 
			pdesc, 
			pvData, 
			&nSize))
	{
		SafeArrayDestroy(psa);
		return DBSTATUS_E_CANTCONVERTVALUE;
	}
	((VARIANTARG *)pbValue)->vt = VT_ARRAY | vt;
	((VARIANTARG *)pbValue)->parray = psa;

	return DBSTATUS_S_OK;
}
