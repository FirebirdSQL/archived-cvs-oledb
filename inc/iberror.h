/*
 * Firebird Open Source OLE DB driver
 *
 * IBERROR.H - Include file for internal error object.
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

#ifndef _IBOLEDB_ERROR_H_
#define _IBOLEDB_ERROR_H_

class CInterBaseError
{
public:
	CInterBaseError();
	~CInterBaseError();
	void ClearError(REFIID);
	HRESULT PostError(WCHAR *, DWORD, HRESULT);
	HRESULT PostError(ISC_STATUS *);
private:
	WCHAR  *m_pwszDesc;
	WCHAR  *m_pwszDescEnd;
	DWORD   m_cbDesc;
	IID     m_riid;
	bool    m_fErrorPosted;
};

#endif /* _IBOLEDB_ERROR_H_ */