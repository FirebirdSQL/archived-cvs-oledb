/*
 * Firebird Open Source OLE DB driver
 *
 * IBACCESSOR.H - Include file for accessor related functions.
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

#ifndef _IBOLEDB_IBACCESSOR_H_
#define _IBOLEDB_IBACCESSOR_H_

typedef struct IBACCESSOR {
	struct IBACCESSOR  *pNext;
	void               *pCreater;
	ULONG               ulSequence;
	DWORD               dwFlags;
	ULONG               ulRefCount;
	ULONG               cBindInfo;
	DWORD               dwParamIO;
	DBBINDING           rgBindings[1];
} IBACCESSOR, *PIBACCESSOR;

#define IBACCESSOR_LENGTH(n) (sizeof(IBACCESSOR) + sizeof(DBBINDING) * ((n) - 1))

HRESULT CreateIbAccessor(
		DBACCESSORFLAGS   dwAccessorFlags,
		ULONG             cBindings,
		const DBBINDING   rgBindings[],
		ULONG             cbRowSize,
		PIBACCESSOR      *ppIbAccessor,
		DBBINDSTATUS      rgStatus[],
		ULONG             cColumns,
		IBCOLUMNMAP      *pIbColumnMap);

HRESULT ReleaseIbAccessor(
		PIBACCESSOR   pIbAccessor, 
		ULONG        *pcRefCount,
		PIBACCESSOR  *ppIbAccessorList);

HRESULT GetIbBindings(
		PIBACCESSOR       pIbAccessor,
		DBACCESSORFLAGS  *pdwAccessorFlags,
		ULONG            *pcBindings,
		DBBINDING       **prgBindings);

HRESULT GetIbColumnInfo(
		ULONG         *pcColumns, 
		DBCOLUMNINFO **prgInfo, 
		OLECHAR      **ppStringsBuffer,
		ULONG          cColumns,
		XSQLDA        *pSqlda,
		IBCOLUMNMAP   *pIbColumnMap);

HRESULT MapIbColumnIDs(
		ULONG         cColumnIDs, 
		const DBID    rgColumnIDs[], 
		ULONG         rgColumns[],
		ULONG         cColumns,
		IBCOLUMNMAP  *pIbColumnMap);

HRESULT SetSqldaPointers(
		XSQLDA       *pSqlda,
		IBCOLUMNMAP  *pIbColumnMap,
		void         *pvData );

#endif /* _IBOLEDB_IBACCESSOR_H_ (not) */