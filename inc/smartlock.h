/*
 * Firebird Open Source OLE DB driver
 *
 * SMARTLOCK.H - Include file for simple syncronization class.
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

#ifndef _IBOLEDB_SMARTLOCK_H_
#define _IBOLEDB_SMARTLOCK_H_

class CSmartLock
{
public:
	CSmartLock::CSmartLock(CRITICAL_SECTION *pcs)
	{
		EnterCriticalSection(pcs);
		m_pcs = pcs;
	}

	CSmartLock::~CSmartLock()
	{
		LeaveCriticalSection(m_pcs);
	}

private:
	CRITICAL_SECTION *m_pcs;
};

#endif /* _IBOLEDB_SMARTLOCK_H_ */