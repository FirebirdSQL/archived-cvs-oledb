/*
 * Firebird Open Source OLE DB driver
 *
 * COMMON.H - Include file for everything that seemed common.
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

#ifndef _IBOLEDB_COMMON_H_
#define _IBOLEDB_COMMON_H_

/* FEATURES */
#define ICOLUMNSROWSET
#define IDBSCHEMAROWSET
#define OPTIONAL_SCHEMA_ROWSETS
//#define IROWSETCHANGE

#ifdef _DEBUG
#define DbgFnTrace2(a) AutoTrace pTrace(a, NULL)
#define DbgFnTrace(a) AutoTrace pTrace(a, this)
#define DbgTrace(a) OutputDebugString(a)
#define DbgPrintGuid(iid) { \
  char szGuid[64]; \
  char *pch; \
  sprintf( szGuid, "%s={%8x-%4x-%4x-%2x%2x-%2x%2x%2x%2x%2x%2x}\n",#iid, iid.Data1, iid.Data2, iid.Data3, iid.Data4[0], iid.Data4[1], iid.Data4[2], iid.Data4[3], iid.Data4[4], iid.Data4[5], iid.Data4[6], iid.Data4[7] ); \
  for(pch=szGuid; *pch!=0;pch++) if(*pch==' ') *pch='0'; \
  OutputDebugString(szGuid); }
#else  /* DEBUG */
#define DbgPrintGuid(iid)
#define DbgFnTrace(a)
#define DbgTrace(a)
#endif /* DEBUG (else) */

#define ALIGN(n, s)		((n + s - 1) & ~(s - 1))
#define ROUNDUP(n, s)	((n + s - 1) & ~(s - 1))

#define PROVIDER_NAME	L"IbOleDb"
#define DEFAULT_COLUMN_COUNT       25
#define MAX_SQL_NAME_SIZE          32
#define STACK_BUFFER_SIZE        1024
#define MAX_SPECIAL_PROPERTYSETS    8
#define IB_NULL       -1
#define IB_DEFAULT    -2
#define IB_NTS        -3

#include <crtdbg.h>
#include <limits.h>
#include <oledb.h>
#include <oledberr.h>
#include <msdadc.h>
#include <atlbase.h>
extern CComModule _Module;
#include <atlcom.h>
#include <crtdbg.h>
#include <ibase.h>
#include "iberror.h"
#include <stdio.h>

typedef DWORD IBHANDLE;
#define IBHANDLE_KEY	(DWORD)4294967291
#ifdef _DEBUG
#define _EncodeHandle(h) ((IBHANDLE)(h))
#define _DecodeAccessorHandle(h) ((PIBACCESSOR)(h))
#define _DecodeRowHandle(h) ((PIBROW)(h))
#else
#define _EncodeHandle(h) ((IBHANDLE)((IBHANDLE)(h)^IBHANDLE_KEY))
#define _DecodeAccessorHandle(h) ((PIBACCESSOR)((IBHANDLE)(h)^IBHANDLE_KEY))
#define _DecodeRowHandle(h) ((PIBROW)((IBHANDLE)(h)^IBHANDLE_KEY))
#endif /* _DEBUG (else) */
/*
 * If we're compiling on a machine that does not have IB 6.0 installed,
 * then we'll need to define this.
 */
#ifndef ISC_TIMESTAMP_DEFINED
#define ISC_TIMESTAMP_DEFINED
typedef ULONG	ISC_TIME;
typedef long	ISC_DATE;
typedef struct {
	ISC_DATE timestamp_date;
	ISC_TIME timestamp_time;} ISC_TIMESTAMP;

#define SQL_TIMESTAMP	SQL_DATE
#define SQL_TYPE_TIME	560
#define SQL_TYPE_DATE	570
#define SQL_INT64		580
#define blr_sql_date	 12
#define blr_sql_time	 13
#define blr_int64		 16

// This is a 6.5 thing, but we'll stick it here anyway
#define DSQL_cancel 4
#endif /* ISC_TIMESTAMP_DEFINED (not) */

typedef struct 
{
	short len;
	char  str[1];
} SQLVARYING;
#define SQLVARYING_LENGTH(l) (sizeof(SQLVARYING) + l)

/*
 * Typedef and extern the functions we'll need to dynamically load if they request
 * SQL Dialect 3 ( IB 6.0 ).  We dynamically load these, so we can still run on 
 * older versions of InterBase.
 */
typedef void (ISC_EXPORT *ISCENCODESQLDATE) (void *, ISC_DATE *);
typedef void (ISC_EXPORT *ISCENCODESQLTIME) (void *, ISC_TIME *);
typedef void (ISC_EXPORT *ISCENCODETIMESTAMP) (void *, ISC_TIMESTAMP *);
typedef void (ISC_EXPORT *ISCDECODESQLDATE) (ISC_DATE *, void *);
typedef void (ISC_EXPORT *ISCDECODESQLTIME) (ISC_TIME *, void *);
typedef void (ISC_EXPORT *ISCDECODETIMESTAMP) (ISC_TIMESTAMP *, void *);

extern ISCENCODESQLDATE g_dso_pIscEncodeSqlDate;
extern ISCENCODESQLTIME g_dso_pIscEncodeSqlTime;
extern ISCENCODETIMESTAMP g_dso_pIscEncodeTimestamp;
extern ISCDECODESQLDATE g_dso_pIscDecodeSqlDate;
extern ISCDECODESQLTIME g_dso_pIscDecodeSqlTime;
extern ISCDECODETIMESTAMP g_dso_pIscDecodeTimestamp;

extern int g_dso_usDialect;

#define strcpy(c1, c2) lstrcpy(c1, c2)
#define wcslen(us1) lstrlenW(us1)
#define stricmp(c1, c2) lstrcmpi(c1, c2)

#define EXPANDGUID(guid) \
	{ guid.Data1, guid.Data2, guid.Data3, \
	{ guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3], guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7] } }

#define _GetStatus(b, d) *((DBSTATUS *)((BYTE *)(d) + (b).obStatus))
#define _GetLength(b, d) (ULONG)*((BYTE *)(d) + (b).obLength)
#define _GetValuePtr(b, d) (void *)((BYTE *)(d) + (b).obValue)

#define BASETYPE(t) ((t) & ~1)
#define NULLABLE(t) ((t) & 1)

#define MIN(a,b) ((a)<(b)?(a):(b))
#define MAX(a,b) ((a)>(b)?(a):(b))

extern HWND			   g_hWnd;
extern HINSTANCE	   g_hInstance;
extern IMalloc		  *g_pIMalloc;
extern IDataConvert	  *g_pIDataConvert;

class CInterBaseDSO;
class CInterBaseSession;
class CInterBaseCommand;
class CInterBaseRowset;

typedef struct {
	ULONG	   ulRefCount;
	DBID	   columnid;
	ULONG	   iOrdinal;
	ULONG	   obData;
	ULONG	   obStatus;
	ULONG	   cbLength;
	ULONG	   ulColumnSize;
	DWORD	   dwFlags;
	DWORD      dwSpecialFlags;
	DBTYPE	   wType;
	BYTE	   bPrecision;
	BYTE	   bScale;
	OLECHAR	   wszAlias[ MAX_SQL_NAME_SIZE ];
	OLECHAR	   wszName[ MAX_SQL_NAME_SIZE ];
	void	  *pvExtra;
} IBCOLUMNMAP;

#define DEFAULT_PAGE_SIZE	4096

typedef struct IBROW {
	struct IBROW  *pNext;
	ULONG          ulRefCount;
	BYTE           pbData[ 1 ];
} IBROW, *PIBROW;

#ifdef _DEBUG
class AutoTrace {
private:
	char	  *m_str;
	void      *m_p;
	char	   m_szMsg[128];

	AutoTrace() { }
	void Begin(char *str, void *p)
	{
		// This is just a pointer copy, but should be good enough...
		m_str = str;
		m_p = p;
		if( m_p != NULL )
			sprintf(m_szMsg, "ENTER %s\n    this=0x%x\n", m_str, m_p);
		else
			sprintf(m_szMsg, "ENTER %s\n", m_str);
		OutputDebugString( m_szMsg ); 
	}

public:

	AutoTrace(char *str, void *p)
	{
		Begin(str, p);
	}

	~AutoTrace()
	{
		if( m_p != NULL )
			sprintf(m_szMsg, "EXIT  %s\n    this=0x%x\n", m_str, m_p);
		else
			sprintf(m_szMsg, "EXIT  %s\n", m_str);
		OutputDebugString( m_szMsg ); 
	}
};
#endif /* _DEBUG */
#endif /* _IBOLEDB_COMMON_H_ (not) */
