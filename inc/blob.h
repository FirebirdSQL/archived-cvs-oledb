/*
 * Firebird Open Source OLEDB driver
 *
 * BLOB.H - Include file for BLOB handling stuff, mainly the
 *          ISequentialStream interface.
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
 * individuals.  Contributors to this file are either listed here or
 * can be obtained from a CVS history command.
 *
 * All rights reserved.
 */

#ifndef _IBOLEDB_BLOB_H_
#define _IBOLEDB_BLOB_H_

class CInterBaseBlobStream : 
	public CComObjectRoot,
	public ISequentialStream
{
public:
	CInterBaseBlobStream();
	~CInterBaseBlobStream();

	Initialize(isc_blob_handle, bool);
	
	/* ISequentialStream methods */
	virtual HRESULT STDMETHODCALLTYPE Read(void *, ULONG, ULONG *);
	virtual HRESULT STDMETHODCALLTYPE Write(void const *, ULONG, ULONG *);

	BEGIN_COM_MAP(CInterBaseBlobStream)
		COM_INTERFACE_ENTRY(ISequentialStream)
	END_COM_MAP()

	DECLARE_GET_CONTROLLING_UNKNOWN()
	DECLARE_NOT_AGGREGATABLE(CInterBaseBlobStream)

private:
	isc_blob_handle   m_iscBlobHandle;
	ISC_STATUS        m_iscStatus[20];
	bool              m_fNullTerminate;
};

DBSTATUS GetArray(
		isc_db_handle *,
		isc_tr_handle *,
		XSQLVAR *,
		BYTE *,
		short *,
		DBTYPE,
		BYTE *,
		DWORD *,
		DWORD *,
		void **);
					
#endif /* _IBOLEDB_BLOB_H_ (not) */