/*
 * Firebird Open Source OLE DB driver
 *
 * UTIL.CPP - Source file for various utility functions.
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

#include <string.h>
#include <memory.h>
#include <windows.h>
#include "../inc/common.h"
#include "../inc/ibaccessor.h"
#include "../inc/rowset.h"
#include "../inc/util.h"
#include "../inc/schemarowset.h"

#define DPB_INITIAL_SIZE	128

DWORD	HashString(
	const WCHAR wszString[])
{
	DWORD	h = 0;
	WCHAR	*pw = (WCHAR *)wszString;

	while ( *pw != L'\0' )
	{
		h = (h << 2) + toupper(*pw);
		pw++;
	}

	return h;
}

DWORD	HashString(
	const char szString[])
{
	DWORD	h = 0;
	char	*p = (char *)szString;

	while ( *p != '\0' )
	{
		h = (h << 2) + toupper(*p);
		p++;
	}

	return h;
}

int dpbinit(char **ppchDPB, short *pcbLen, char chPBType)
{
	char *pchDPB = *ppchDPB;

	*pcbLen = 1;

	if (!pchDPB)
	{
		pchDPB = new char[DPB_INITIAL_SIZE + sizeof(short)];
		*(short *)pchDPB = DPB_INITIAL_SIZE;
		pchDPB += sizeof(short);
	}

	*pchDPB = chPBType;
	*ppchDPB = pchDPB;

	return 1;
}

int dpbcat(char **ppchDPB, short *pcbLen, char chItem, void *pvValue)
{
	char *pchDPB = *ppchDPB;
	short cbDPB = *pcbLen;
	char *pchNext = pchDPB + cbDPB;
	short *pcbBufLen = (short *)(pchDPB - sizeof(short));
	short cbItemLen;

	switch (chItem)
	{
	case isc_dpb_user_name:
	case isc_dpb_password:
	case isc_dpb_sql_role_name:
	case isc_dpb_lc_ctype:
		cbItemLen = lstrlenA((char *)pvValue) + 1;
		break;
	default:
		cbItemLen = ( pvValue == NULL ? 0 : 2 );
	}

	cbDPB += 1 + cbItemLen;

	if (*pcbBufLen < cbDPB)
	{
		char *pchNewDPB = new char[cbDPB + sizeof(short)];
		*pcbBufLen = cbDPB;

		memcpy(pchNewDPB, pchDPB, *pcbLen + sizeof(short));
		delete[] pcbBufLen;
		pchDPB = pchNewDPB + sizeof(short);
		pchNext = pchDPB + *pcbLen;
		*ppchDPB = pchDPB;
	}

	*pchNext++ = chItem;

	if (pvValue)
	{
		*pchNext++ = (char)cbItemLen - 1;
		memcpy(pchNext, pvValue, cbItemLen);
	}

	*pcbLen = cbDPB;

	return *pcbLen;
}

int dpbfree(char **ppDPB, short *pcbLen)
{
	char *pDPB = *ppDPB;
	pDPB -= sizeof(short);
	delete[] pDPB;
	*ppDPB = 0;
	pcbLen = 0;
	return 0;
}