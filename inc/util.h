/*
 * Firebird Open Source OLE DB driver
 *
 * UTIL.H - Include file for various utility functions.
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

#ifndef _IBOLEDB_UTIL_H_
#define _IBOLEDB_UTIL_H_

int sqlnamecmp(
		char	  *pszName1,
		char	  *pszName2,
		int		   nLen);

int dpbinit(char **, short *, char);
int dpbcat(char **, short *, char, void *);
int dpbfree(char **, short *);

DWORD	HashString(const WCHAR	wszString[] );
DWORD	HashString(const char	szString[] );

HRESULT GetKeyColumnInfo(
		XSQLDA         *sqlda,
		IBCOLUMNMAP    *pColumnMap,
		isc_db_handle  *piscDbHandle,
		isc_tr_handle  *piscTransHandle);

#endif /* _IBOLEDB_UTIL_H_ */