/*
 * Firebird Open Source OLE DB driver
 *
 * CONVERT.CPP - Source file for OLE DB data type conversion functions.
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

#include <time.h>
#include "../inc/common.h"
#include "../inc/convert.h"
#include "../inc/util.h"

static int StringDivN
	(
	char		  *pszNumber,
	int			   nDiv
	);
static DBSTATUS StringToVal
	(
	char		  *pszNumber,
	BYTE		   val[]
	);

DBSTATUS ClosestOleDbType(
	short    sqltype,
	short    sqlsubtype,
	short    sqllen,
	short    sqlscale,
	DBTYPE  *pwType,
	DWORD   *pdwSize,
	BYTE    *pbPrecision,
	BYTE    *pbScale,
	DWORD   *pdwFlags)
{
	*pbScale = -sqlscale;
	*pbPrecision = 0;
	*pdwSize = sqllen;
	*pdwFlags = DBCOLUMNFLAGS_ISFIXEDLENGTH;

	switch( BASETYPE( sqltype ) )
	{
	case SQL_TEXT:
		*pwType = DBTYPE_STR;
		return DBSTATUS_S_OK;

	case SQL_VARYING:
		//*pdwSize -= sizeof(short);
		*pdwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
		*pwType = DBTYPE_STR;
		return DBSTATUS_S_OK;

	case SQL_SHORT:
		if (sqlscale)
		{
			*pbPrecision = 4;
			*pdwSize = sizeof(DB_NUMERIC);
			*pwType = DBTYPE_NUMERIC;
		}
		else
			*pwType = DBTYPE_I2;
		return DBSTATUS_S_OK;

	case SQL_LONG:
		if (sqlscale)
		{
			*pbPrecision = 9;
			*pdwSize = sizeof(DB_NUMERIC);
			*pwType = DBTYPE_NUMERIC;
		}
		else
			*pwType = DBTYPE_I4;
		return DBSTATUS_S_OK;

	case SQL_FLOAT:
		*pwType = DBTYPE_R4;
		return DBSTATUS_S_OK;

	case SQL_DOUBLE:
#ifdef DOUBLE_AS_NUMERIC
		if (sqlscale)
		{
			*pbPrecision = 15;
			*pdwSize = sizeof(DB_NUMERIC);
			*pwType = DBTYPE_NUMERIC;
		}
		else
#endif /* DOUBLE_AS_NUMERIC */
			*pwType = DBTYPE_R8;
		return DBSTATUS_S_OK;

	case SQL_TIMESTAMP: /* SQL_DATE */
		*pdwSize = sizeof(DBTIMESTAMP);
		*pwType = DBTYPE_DBTIMESTAMP;
		return DBSTATUS_S_OK;

	case SQL_BLOB:
		*pdwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
		*pdwFlags |= DBCOLUMNFLAGS_ISLONG;
		*pdwSize = ~0;
		*pwType = sqlsubtype == isc_blob_text ? DBTYPE_STR : DBTYPE_BYTES;
		return DBSTATUS_S_OK;

	case SQL_ARRAY:
		*pdwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
		*pdwFlags |= DBCOLUMNFLAGS_ISLONG;
		*pdwSize = -1;
		*pwType = DBTYPE_ARRAY;
		return DBSTATUS_S_OK;

	case SQL_QUAD:
		return DBSTATUS_S_OK;

	case SQL_TYPE_DATE:
		*pdwSize = sizeof(DBDATE);
		*pwType = DBTYPE_DBDATE;
		return DBSTATUS_S_OK;

	case SQL_TYPE_TIME:
		*pdwSize = sizeof(DBTIME);
		*pwType = DBTYPE_DBTIME;
		return DBSTATUS_S_OK;

	case SQL_INT64:
		if (sqlscale)
		{
			*pbPrecision = 18;
			*pdwSize = sizeof(DB_NUMERIC);
			*pwType = DBTYPE_NUMERIC;
		}
		else
			*pwType = DBTYPE_I8;

		return DBSTATUS_S_OK;
	}
	return DBTYPE_EMPTY;
}

DBSTATUS InterBaseFromOleDbType(
		short	   SqlType,
		short	   SqlSubtype,
		short	   SqlLen,
		short	   SqlScale,
		BYTE	  *pSqlData,
		short	  *pSqlInd,
		DBTYPE	   wType,
		DWORD	   dwSize,
		BYTE	   bPrecision,
		BYTE	   bScale,
		void	  *pvData,
		DBSTATUS   wStatus)
{
	DBTYPE		   wClosestOleDbType;
	DWORD		   dwFlags;
	DWORD		   dwClosestSize;
	BYTE		  *pbBuffer = (BYTE *)pvData;
	BYTE		  *pbBufferToFree = NULL;

	if (wStatus & DBSTATUS_S_ISNULL)
	{
		*pSqlInd = (short)-1;
		return wStatus;
	}

	if ((wStatus = ClosestOleDbType(
			SqlType, 
			SqlSubtype, 
			SqlLen, 
			SqlScale, 
			&wClosestOleDbType, 
			&dwClosestSize,
			&bPrecision, 
			&bScale, 
			&dwFlags)) != DBSTATUS_S_OK)
		return DBSTATUS_E_CANTCONVERTVALUE;

#ifdef DOUBLE_AS_NUMERIC
	if (wClosestOleDbType == DBTYPE_NUMERIC)
	{
		wClosestOleDbType = DBTYPE_STR;
		dwClosestSize = 38;
	}
#endif

	if (wClosestOleDbType != wType)
	{
		DBSTATUS	   wTempStatus;

		if( wClosestOleDbType == DBTYPE_STR )
		{
			dwClosestSize++;
		}

		if( ( dwClosestSize ) <= STACK_BUFFER_SIZE )
		{
			pbBuffer = (BYTE *)alloca( dwClosestSize );
		}
		else
		{
			pbBuffer = new BYTE[dwClosestSize];
			pbBufferToFree = pbBuffer;
		}

		if (FAILED(g_pIDataConvert->DataConvert(
				wType,				//wSrcType
				wClosestOleDbType,	//wDstType
				dwSize,				//cbSrcLength
				&dwClosestSize,		//pcbDstLength
				pvData,				//pSrc
				pbBuffer,			//pDst
				dwClosestSize,		//pDstMaxLength
				wStatus,			//dbsSrcStatus
				&wTempStatus,		//pdbsStatus
				bPrecision,			//bPrecision
				bScale,				//bScale
				dwFlags)))			//dwFlags
			return DBSTATUS_E_CANTCONVERTVALUE;

	//	pvClosestData = pbBuffer;
	}

	switch( BASETYPE( SqlType ) )
	{
	case SQL_VARYING:
		*(short *)pSqlData = lstrlenA((char *)pbBuffer);
		strcpy((char *)pSqlData + sizeof(short), (char *)pbBuffer);
		break;
	
	case SQL_TYPE_DATE:
		if( g_dso_pIscEncodeSqlDate != NULL )
		{
			struct tm tmDate;

			memset( &tmDate, 0, sizeof(tmDate) );
			tmDate.tm_year = ((DBDATE *)pbBuffer)->year - 1900;
			tmDate.tm_mon  = ((DBDATE *)pbBuffer)->month - 1;
			tmDate.tm_mday = ((DBDATE *)pbBuffer)->day;

			g_dso_pIscEncodeSqlDate( &tmDate, (ISC_DATE *)pSqlData );

			dwSize = sizeof(DBDATE);
		}
		else
			return DBSTATUS_E_CANTCONVERTVALUE;
		break;
	
	case SQL_TYPE_TIME:
		if( g_dso_pIscEncodeSqlTime != NULL )
		{
			struct tm tmDate;

			memset( &tmDate, 0, sizeof(tmDate) );
			tmDate.tm_hour = ((DBTIME *)pbBuffer)->hour;
			tmDate.tm_min  = ((DBTIME *)pbBuffer)->minute;
			tmDate.tm_sec  = ((DBTIME *)pbBuffer)->second;

			g_dso_pIscEncodeSqlTime( &tmDate, (ISC_TIME *)pSqlData );

			dwSize = sizeof(DBTIME);
		}
		else
			return DBSTATUS_E_CANTCONVERTVALUE;
		break;

	case SQL_TIMESTAMP: /* SQL_DATE */
		{
			struct tm tmDate;

			memset( &tmDate, 0, sizeof(tmDate) );

			tmDate.tm_year = ((DBTIMESTAMP *)pbBuffer)->year - 1900;
			tmDate.tm_mon  = ((DBTIMESTAMP *)pbBuffer)->month - 1;
			tmDate.tm_mday = ((DBTIMESTAMP *)pbBuffer)->day;
			tmDate.tm_hour = ((DBTIMESTAMP *)pbBuffer)->hour;
			tmDate.tm_min  = ((DBTIMESTAMP *)pbBuffer)->minute;
			tmDate.tm_sec  = ((DBTIMESTAMP *)pbBuffer)->second;

			if( g_dso_pIscEncodeTimestamp != NULL )
				g_dso_pIscEncodeTimestamp( &tmDate, (ISC_TIMESTAMP *)pSqlData );
			else
				isc_encode_date( &tmDate, (ISC_QUAD *)pSqlData );

			dwSize = sizeof(DBTIMESTAMP);
		}
		break;

	case SQL_BLOB:
		break;
	
	case SQL_DOUBLE:
#ifdef DOUBLE_AS_NUMERIC
		if (SqlScale)
		{
			*(double *)pSqlData = atof((char *)pbBuffer);
			break;
		}
#endif
		*(double *)pSqlData = *(double *)pbBuffer;
		break;

	case SQL_SHORT:
		if( wClosestOleDbType == DBTYPE_NUMERIC )
		{
			if( ( (DB_NUMERIC *)pbBuffer )->sign == 0 )
				*(USHORT *)pSqlData = ~( ( *(USHORT *)( (DB_NUMERIC *)pbBuffer )->val ) - 1 );
			else
				*(USHORT *)pSqlData = *(USHORT *)( (DB_NUMERIC *)pbBuffer )->val;
		}
		else
		{
			*(SHORT *)pSqlData = *(SHORT *)pbBuffer;
		}
		break;

	case SQL_LONG:
		if( wClosestOleDbType == DBTYPE_NUMERIC )
		{
			if( ( (DB_NUMERIC *)pbBuffer )->sign == 0 )
				*(ULONG *)pSqlData = ~( ( *(ULONG *)( (DB_NUMERIC *)pbBuffer )->val ) - 1 );
			else
				*(ULONG *)pSqlData = *(ULONG *)( (DB_NUMERIC *)pbBuffer )->val;
		}
		else
		{
			*(long *)pSqlData = *(long *)pbBuffer;
		}
		break;

	case SQL_INT64:
		if( wClosestOleDbType == DBTYPE_NUMERIC )
		{
			if( ( (DB_NUMERIC *)pbBuffer )->sign == 0 )
				*(ULONGLONG *)pSqlData = ~( ( *(ULONGLONG *)( (DB_NUMERIC *)pbBuffer )->val ) - 1 );
			else
				*(ULONGLONG *)pSqlData = *(ULONGLONG *)( (DB_NUMERIC *)pbBuffer )->val;
		}
		else
		{
			*(LONGLONG *)pSqlData = *(LONGLONG *)pbBuffer;
		}
		break;

	case SQL_FLOAT:
	case SQL_TEXT:
		_ASSERTE(dwClosestSize == (DWORD)SqlLen);
		memcpy(pSqlData, pbBuffer, SqlLen);
		break;
	default:
		break;
	}

	if (pbBufferToFree != NULL)
		delete[] pbBufferToFree;

	return DBSTATUS_S_OK;
}

DBSTATUS InterBaseToOleDbType(
//		XSQLVAR   *pSqlVar,
		short	   sqltype,
		short	   sqlsubtype,
		short	   sqllen,
		short	   sqlscale,
		BYTE	  *sqldata,
		short	  *sqlind,
		DBTYPE	   wType,
		DWORD	  *pdwSize,
		BYTE	  *pbPrecision,
		BYTE	  *pbScale,
		void	  *pvData,
		DBSTATUS  *pwStatus
		)
{
	DBSTATUS   wStatus;
	DWORD	   dwFlags;
	DWORD	   dwSize;
	DBTYPE	   wClosestOleDbType;
	BYTE	  *pbValue = sqldata;
	union {
		DB_NUMERIC	   dbNumeric;
		DBTIMESTAMP	   dbTimeStamp;
		DBDATE		   dbDate;
		DBTIME		   dbTime;
	} dbTemp;
	struct tm  tmDate;
	bool       fByRef = false;

	if (!(pwStatus && pdwSize && pvData))
	{
		return E_INVALIDARG;
	}

	if ((wStatus = ClosestOleDbType(
			sqltype, 
			sqlsubtype, 
			sqllen, 
			sqlscale, 
			&wClosestOleDbType, 
			&dwSize,
			pbPrecision, 
			pbScale, 
			&dwFlags)) != DBSTATUS_S_OK)
	{
		return DBSTATUS_E_CANTCONVERTVALUE;
	}

	if( NULLABLE( sqltype ) && *sqlind )
	{
		wStatus = DBSTATUS_S_ISNULL;
	}
	else
	{
		*pbScale = (BYTE)-sqlscale;
		*pbPrecision = 0;

		switch( BASETYPE( sqltype ) )
		{
		case SQL_VARYING:
			dwSize = *(short *)pbValue;
			pbValue += sizeof(short);
			break;
			
		case SQL_SHORT:
			if( sqlscale != 0 )
			{
				dbTemp.dbNumeric.sign = *(short *)sqldata > 0;
				dbTemp.dbNumeric.scale = -sqlscale;
				dbTemp.dbNumeric.precision = 4;
				memset(dbTemp.dbNumeric.val, 0, sizeof(dbTemp.dbNumeric.val));
				if( dbTemp.dbNumeric.sign == 0 )
				{
					*(USHORT *)dbTemp.dbNumeric.val = ~( *(USHORT *)sqldata ) + 1;
				}
				else
				{
					*(USHORT *)dbTemp.dbNumeric.val = *(USHORT *)sqldata;
				}
				pbValue = (BYTE *)&(dbTemp.dbNumeric);
				dwSize = sizeof(DB_NUMERIC);
			}
			break;

		case SQL_LONG:
			if( sqlscale != 0 )
			{
				dbTemp.dbNumeric.sign = *(long *)sqldata > 0;
				dbTemp.dbNumeric.scale = -sqlscale;
				dbTemp.dbNumeric.precision = 9;
				memset(dbTemp.dbNumeric.val, 0, sizeof(dbTemp.dbNumeric.val));
				if( dbTemp.dbNumeric.sign == 0 )
				{
					*(ULONG *)dbTemp.dbNumeric.val = ~( *(ULONG *)sqldata ) + 1;
				}
				else
				{
					*(ULONG *)dbTemp.dbNumeric.val = *(ULONG *)sqldata;
				}
				pbValue = (BYTE *)&(dbTemp.dbNumeric);
				dwSize = sizeof(DB_NUMERIC);
			}
			break;

		case SQL_INT64:
			if( sqlscale != 0 )
			{
				dbTemp.dbNumeric.sign = *(LONGLONG *)sqldata > 0;
				dbTemp.dbNumeric.scale = -sqlscale;
				dbTemp.dbNumeric.precision = 18;
				memset(dbTemp.dbNumeric.val, 0, sizeof(dbTemp.dbNumeric.val));
				if( dbTemp.dbNumeric.sign == 0 )
				{
					*(ULONGLONG *)dbTemp.dbNumeric.val = ~( *(ULONGLONG *)sqldata ) + 1;
				}
				else
				{
					*(ULONGLONG *)dbTemp.dbNumeric.val = *(ULONGLONG *)sqldata;
				}
				pbValue = (BYTE *)&(dbTemp.dbNumeric);
				dwSize = sizeof(DB_NUMERIC);
			}
			break;

		case SQL_TEXT:
		case SQL_FLOAT:
			break;
			
		case SQL_DOUBLE:
#ifdef DOUBLE_AS_NUMERIC
			if( sqlscale != 0 )
			{
				int			   nSign;
				int			   nDecimal;
				int			   nScale = -sqlscale;
				char		  *szDigits;

				szDigits = _fcvt(*(double *)sqldata, nScale + 1, &nDecimal, &nSign );
				szDigits[lstrlenA(szDigits) - 1] = '\0';

				dbTemp.dbNumeric.sign = !nSign;
				dbTemp.dbNumeric.scale = nScale;
				dbTemp.dbNumeric.precision = 15;	//strlen(szDigits);
				StringToVal(szDigits, dbTemp.dbNumeric.val);
				pbValue = (BYTE *)&(dbTemp.dbNumeric);
				dwSize = sizeof(DB_NUMERIC);
			}
#endif
			break;
			
		case SQL_D_FLOAT:
			wStatus = DBSTATUS_E_CANTCONVERTVALUE;
			break;
		
		case SQL_TYPE_DATE:
			if( g_dso_pIscDecodeSqlDate != NULL )
			{
				g_dso_pIscDecodeSqlDate( (ISC_DATE *)sqldata, &tmDate );
				dwSize = sizeof(DBDATE);
				dbTemp.dbDate.day = tmDate.tm_mday;
				dbTemp.dbDate.month = 1 + tmDate.tm_mon;
				dbTemp.dbDate.year = 1900 + tmDate.tm_year;
				pbValue = (BYTE *)&(dbTemp.dbDate);
			}
			else
			{
				wStatus = DBSTATUS_E_CANTCONVERTVALUE;
			}
			break;

		case SQL_TYPE_TIME:
			if( g_dso_pIscDecodeSqlTime != NULL )
			{
				g_dso_pIscDecodeSqlTime( (ISC_TIME *)sqldata, &tmDate );
				dwSize = sizeof(DBTIME);
				dbTemp.dbTime.hour = tmDate.tm_hour;
				dbTemp.dbTime.minute = tmDate.tm_min;
				dbTemp.dbTime.second = tmDate.tm_sec;
				pbValue = (BYTE *)&(dbTemp.dbTime);
			}
			else
			{
				wStatus = DBSTATUS_E_CANTCONVERTVALUE;
			}
			break;

		case SQL_TIMESTAMP:	/* SQL_DATE */
			if( g_dso_pIscDecodeTimestamp != NULL )
			{
				g_dso_pIscDecodeTimestamp( (ISC_TIMESTAMP *)sqldata, &tmDate );
			}
			else
			{
				isc_decode_date( (ISC_QUAD *)sqldata, &tmDate );
			}

			dwSize = sizeof(DBTIMESTAMP);
			dbTemp.dbTimeStamp.year = 1900 + tmDate.tm_year;
			dbTemp.dbTimeStamp.month = 1 + tmDate.tm_mon;
			dbTemp.dbTimeStamp.day = tmDate.tm_mday;
			dbTemp.dbTimeStamp.hour = tmDate.tm_hour;
			dbTemp.dbTimeStamp.minute = tmDate.tm_min;
			dbTemp.dbTimeStamp.second = tmDate.tm_sec;
			dbTemp.dbTimeStamp.fraction = 0;
			pbValue = (BYTE *)&(dbTemp.dbTimeStamp);
			break;
			
		case SQL_BLOB:
			wStatus = DBSTATUS_E_CANTCONVERTVALUE;
			break;

		case SQL_ARRAY:
			wStatus = DBSTATUS_E_CANTCONVERTVALUE;
			break;

		case SQL_QUAD:
			break;
		}
	}

	if( (*pwStatus = wStatus) == DBSTATUS_S_OK )
	{
		if (wType == wClosestOleDbType)
		{
			void *pvBuffer = pvData;

			if( fByRef )
			{
				pvBuffer = g_pIMalloc->Alloc( dwSize + 1 );
				*( (void **)pvData ) = pvBuffer;
			}
			memcpy(pvBuffer, pbValue, dwSize);
			if (wType == DBTYPE_STR)
			{
				((char *)pvBuffer)[dwSize] = '\0';
			}
			*pdwSize = dwSize;
		}
		else
		{
			if( FAILED( g_pIDataConvert->DataConvert(
					wClosestOleDbType,
					wType,
					dwSize,
					pdwSize,
					pbValue,
					pvData,
					*pdwSize,
					wStatus,
					pwStatus,
					*pbPrecision,
					*pbScale,
					dwFlags) ) )
			{
				dwFlags = DBSTATUS_E_CANTCONVERTVALUE;
			}
		}
	}

	return wStatus;
}

HRESULT GetStringFromVariant(
		char		 **ppchString, 
		ULONG		   cbString,
		ULONG		  *pcbBytesUsed,
		VARIANTARG	  *pv )
{
	char      *pchString = *ppchString;
	ULONG	   nSize;
	VARIANT    tTemp[ 1 ];

	VariantInit( tTemp );

	if( pchString == NULL &&
		cbString != 0 )
	{
		return E_INVALIDARG;
	}

	if( V_VT( pv ) != VT_BSTR )
	{
		if( VariantChangeType( tTemp, pv, 0, VT_BSTR ) != S_OK )
		{
			return E_INVALIDARG;
		}
		pv = tTemp;
	}

	nSize = WideCharToMultiByte(
			CP_ACP, 
			0,
			V_BSTR( pv ),
			-1,
			pchString,
			cbString,
			NULL,
			NULL);

	if( pcbBytesUsed != NULL )
	{
		*pcbBytesUsed = nSize - 1;
	}

	if( cbString != 0 )
	{
		if( nSize > 0 )
		{
			return S_OK;
		}

		return E_FAIL;
	}

	if( ( pchString = new char[ nSize ] ) == NULL )
	{
		return E_OUTOFMEMORY;
	}

	WideCharToMultiByte(
			CP_ACP, 
			0,
			V_BSTR( pv ),
			-1,
			pchString,
			nSize,
			NULL,
			NULL );

	VariantClear( tTemp );

	*ppchString = pchString;

	return S_OK;
}

static DBSTATUS StringToVal
	(
	char		  *pszNumber,
	BYTE		   val[]
	)
{
	char	  *pchPtr;
	char	   szNum[39 + 1];
	BYTE	   rgBytes[16];
	BYTE	  *pbCurrentByte;
	BYTE	   bPrecision;
	ULONG	   cbSize;
	
	bPrecision = lstrlenA(pszNumber);

	strcpy(szNum, pszNumber);
	pbCurrentByte = rgBytes;

	while ( 1 )
	{
		*pbCurrentByte++ = (char)StringDivN(szNum, 0x100);

		for (pchPtr = szNum; *pchPtr != '\0'; pchPtr++)
			if (*pchPtr != '0')
				break;
		if (*pchPtr == '\0')
			break;
	}

	cbSize = pbCurrentByte - rgBytes;
	memset(val, 0, sizeof(rgBytes));
//	memcpy(val + sizeof(rgBytes) - cbSize, pbCurrentByte, cbSize);
	memcpy(val, rgBytes, cbSize);
	return DBSTATUS_S_OK;
}

static int StringDivN
	(
	char		  *pszNumber,
	int			   nDiv
	)
{
	int	   n;
	int	   r = 0;
	char	  *p;

	for (p = pszNumber; *p; p++)
	{
		n = *p - '0';
		n = n + r * 10;
		r = n % nDiv;
		n = n / nDiv;
		*p = n + '0';
	}
	return r;
}

DBTYPE BlrTypeToOledbType( 
	short sType,
	short sSubType,
	DBTYPE *pwType,
	USHORT *pusPrecision,
	SHORT  *psScale,
	DWORD  *pdwFlags,
	ULONG  *pulCharLength,
	ULONG  *pulByteLength )
{
	*pdwFlags = DBCOLUMNFLAGS_ISFIXEDLENGTH;

	switch ( sType )
	{
	case blr_varying:   //VARYING
		*pdwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
	case blr_text:      //TEXT
		*pwType = DBTYPE_STR;
		break;

	case blr_blob:
		*pulByteLength = *pulCharLength = 0;
		*pdwFlags &= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
		*pdwFlags |= DBCOLUMNFLAGS_ISLONG;
		switch( sSubType )
		{
		case isc_blob_text:
			*pwType = DBTYPE_STR;
			break;
		default:
			*pwType = DBTYPE_BYTES;
		}
		break;

	case blr_short:
		if( *psScale == 0 )
		{
			*pulCharLength = 5;
			*pwType = DBTYPE_I2;
		}
		else
		{
			*pulCharLength = *pusPrecision = 4;
			*psScale = -*psScale;
			*pwType = DBTYPE_NUMERIC;
		}
		break;

	case blr_long:
		if( *psScale == 0 )
		{
			*pulCharLength = 10;
			*pwType = DBTYPE_I4;
		}
		else
		{
			*pulCharLength = *pusPrecision = 9;
			*psScale = -*psScale;
			*pwType = DBTYPE_NUMERIC;
		}
		break;

	case blr_float:
		*pulCharLength = 7;
		*pwType = DBTYPE_R4;
		break;

	case blr_double:
		*pulCharLength = 15;
		*pwType = DBTYPE_R8;
		break;

	case blr_int64:
		if( *psScale == 0 )
		{
			*pulCharLength = 19;
			*pwType = DBTYPE_I8;
		}
		else
		{
			*pulCharLength = *pusPrecision = 18;
			*psScale = -*psScale;
			*pwType = DBTYPE_NUMERIC;
		}
		break;

	case blr_sql_time:
		*pulCharLength = 8;
		*pusPrecision = 0;
		*pwType = DBTYPE_DBTIME;
		break;

	case blr_sql_date:
		*pulCharLength = 10;
				*pusPrecision = 0;
		*pwType = DBTYPE_DBDATE;
		break;

	case blr_date:
		*pulCharLength = 23;
		*pusPrecision = 3;
		*pwType = DBTYPE_DBTIMESTAMP;

	}
	return 0;
}
