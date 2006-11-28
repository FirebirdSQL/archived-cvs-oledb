/*
 * Firebird Open Source OLE DB driver
 *
 * CONVERT.H - Include file for OLE DB data type conversion functions.
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

#ifndef _IBOLEDB_CONVERT_H_
#define _IBOLEDB_CONVERT_H_

#define NULL_STR_VALUE		"<NULL>"
#define NULL_STR_VALUE_SIZE	7
#define NULL_WSTR_VALUE		L"<NULL>"

#ifndef MIN
#define MIN(a,b) ((a)<(b)?(a):(b))
#endif

DBSTATUS ClosestOleDbType(
		short	   sqltype,
		short	   sqlsubtype,
		short	   sqllen,
		short	   sqlscale,
		DBTYPE	  *pwType,
		DWORD	  *pdwSize,
		BYTE	  *pbPrecision,
		BYTE	  *pbScale,
		DWORD	  *pdwFlags);

DBSTATUS InterBaseFromOleDbType(
		short	   sqltype,
		short	   sqlsubtype,
		short	   sqllen,
		short	   sqlscale,
		BYTE	  *sqldata,
		short	  *sqlind,
		DBTYPE	   wType,
		DWORD	   dwSize,
		BYTE	   bPrecision,
		BYTE	   bScale,
		void	  *pvData,
		DBSTATUS   wStatus);

DBSTATUS InterBaseToOleDbType(
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
		DBSTATUS  *pwStatus);

HRESULT GetStringFromVariant(
		char       **ppchString, 
		ULONG        cbString, 
		ULONG       *pcbBytesUsed,
		VARIANTARG  *pv );

DBTYPE BlrTypeToOledbType( 
	short    sType,
	short    sSubType,
	DBTYPE  *psType,
	USHORT  *pusPrecision,
	SHORT   *psScale,
	DWORD   *pdwFlags,
	ULONG   *pulCharLength,
	ULONG   *pulByteLength );
#endif /* _IBOLEDB_CONVERT_H_ (not) */