/*
 * Firebird Open Source OLE DB driver
 *
 * ROWSET.H - Include file for OLE DB rowset object definition
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

#ifndef _IBOLEDB_ROWSET_H_
#define _IBOLEDB_ROWSET_H_

#include "ibaccessor.h"

class CInterBaseRowset :
	public CComObjectRoot,
	public IAccessor,
	public IColumnsInfo,
#ifdef ICOLUMNSROWSET
	public IColumnsRowset,
#endif /* ICOLUMNSROWSET */
	public IConvertType,
	public IRowset,
#ifdef IROWSETCHANGE
	public IRowsetChange,
#endif /* IROWSETCHANGE */
#ifdef IROWSETIDENTITY
	public IRowsetIdentity,
#endif /* IROWSETIDENTITY */
	public IRowsetInfo,
	public ISupportErrorInfo
{
public:
	CInterBaseRowset();
	~CInterBaseRowset();

	void  *AllocPage( ULONG *pulPageSize );
	PIBROW AllocRow( void );
	HRESULT InsertPrefetchedRow( PIBROW );
	HRESULT SetPrefetchedColumnInfo(const PrefetchedColumnInfo [], ULONG, IBCOLUMNMAP **);
	HRESULT SetPrefetchedColumnData(ULONG, PIBROW, short, short, void *);

	HRESULT Initialize(
			CInterBaseCommand *, 
			CInterBaseSession *, 
			isc_db_handle *,
			isc_tr_handle *,
			isc_stmt_handle,
			ULONG        cbRowBufferSize,
			XSQLDA      *pOutSqlda,
			IBCOLUMNMAP *pIbColumnMap,
			XSQLDA      *pInSqlda,
			IBCOLUMNMAP *pIbParamMap,
			ULONG        ulAccessorSequence,
			bool         fUserCommand );

	/* The ATL interface map */
	BEGIN_COM_MAP(CInterBaseRowset)
		COM_INTERFACE_ENTRY(IAccessor)
		COM_INTERFACE_ENTRY(IColumnsInfo)
#ifdef ICOLUMNSROWSET
		COM_INTERFACE_ENTRY(IColumnsRowset)
#endif /*ICOLUMNSROWSET */
		COM_INTERFACE_ENTRY(IConvertType)
		COM_INTERFACE_ENTRY(IRowset)
#ifdef IROWSETCHANGE
		COM_INTERFACE_ENTRY(IRowsetChange)
#endif /* IROWSETCHANGE */
#ifdef IROWSETIDENTITY
		COM_INTERFACE_ENTRY(IRowsetIdentity)
#endif /* IROWSETIDENTITY */
		COM_INTERFACE_ENTRY(IRowsetInfo)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
	END_COM_MAP()

	DECLARE_GET_CONTROLLING_UNKNOWN()
	DECLARE_AGGREGATABLE(CInterBaseRowset)

/* IAccessor */
virtual HRESULT STDMETHODCALLTYPE AddRefAccessor(HACCESSOR, ULONG *);
virtual HRESULT STDMETHODCALLTYPE CreateAccessor(DBACCESSORFLAGS, ULONG, const DBBINDING [], ULONG, HACCESSOR *, DBBINDSTATUS []);
virtual HRESULT STDMETHODCALLTYPE GetBindings(HACCESSOR, DBACCESSORFLAGS *, ULONG *,DBBINDING ** );
virtual HRESULT STDMETHODCALLTYPE ReleaseAccessor(HACCESSOR, ULONG *);

/* IColumnsInfo */
virtual HRESULT STDMETHODCALLTYPE GetColumnInfo(ULONG *, DBCOLUMNINFO ** , OLECHAR ** );
virtual HRESULT STDMETHODCALLTYPE MapColumnIDs(ULONG, const DBID [], ULONG []);

#ifdef ICOLUMNSROWSET
/* IColumnsRowset */
virtual HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetAvailableColumns(ULONG *, DBID **);
virtual HRESULT STDMETHODCALLTYPE CInterBaseRowset::GetColumnsRowset(IUnknown *, ULONG, const DBID [], REFIID, ULONG, DBPROPSET [], IUnknown **);
#endif /* ICOLUMNSROWSET */

//IConvertType
virtual HRESULT STDMETHODCALLTYPE CanConvert(DBTYPE, DBTYPE, DBCONVERTFLAGS);

/* IRowset */
virtual HRESULT STDMETHODCALLTYPE AddRefRows(ULONG, const HROW [], ULONG [], DBROWSTATUS []);
virtual HRESULT STDMETHODCALLTYPE GetData(HROW, HACCESSOR, void *);
virtual HRESULT STDMETHODCALLTYPE GetNextRows(HCHAPTER, LONG, LONG, ULONG *, HROW ** );
virtual HRESULT STDMETHODCALLTYPE ReleaseRows(ULONG, const HROW [], DBROWOPTIONS [],ULONG [], DBROWSTATUS []);
virtual HRESULT STDMETHODCALLTYPE RestartPosition(HCHAPTER);

/* IRowsetChange */
#ifdef IROWSETCHANGE
virtual HRESULT STDMETHODCALLTYPE DeleteRows( HCHAPTER, ULONG,const HROW [],DBROWSTATUS []);
virtual HRESULT STDMETHODCALLTYPE InsertRow( HCHAPTER, HACCESSOR, void *, HROW *);
virtual HRESULT STDMETHODCALLTYPE SetData( HROW, HACCESSOR, void *);
#endif //IROWSETCHANGE

#ifdef IROWSETIDENTITY
/* RowsetIdentity */
virtual HRESULT STDMETHODCALLTYPE IsSameRow(HROW, HROW);
#endif /* IROWSETIDENTITY */

/* IRowsetInfo */
virtual HRESULT STDMETHODCALLTYPE GetProperties(const ULONG, const DBPROPIDSET [], ULONG *, DBPROPSET ** );
virtual HRESULT STDMETHODCALLTYPE GetReferencedRowset(ULONG, REFIID, IUnknown ** );
virtual HRESULT STDMETHODCALLTYPE GetSpecification(REFIID, IUnknown ** );

/* ISupportErrorInfo */
STDMETHODIMP InterfaceSupportsErrorInfo(REFIID iid);

private:
	CRITICAL_SECTION    m_cs;
	ISC_STATUS          m_iscStatus[20];
	CInterBaseError    *m_pInterBaseError;
	CInterBaseSession  *m_pOwningSession;
	CInterBaseCommand  *m_pOwningCommand;
	XSQLDA             *m_pInSqlda;
	XSQLDA             *m_pOutSqlda;
	isc_db_handle      *m_piscDbHandle;
	isc_tr_handle      *m_piscTransHandle;
	IBACCESSOR         *m_pIbAccessorsToRelease;
	IBCOLUMNMAP        *m_pIbColumnMap;
	IBROW              *m_pFirstPrefetchedIbRow;
	IBROW              *m_pCurrentPrefetchedIbRow;
	IBROW              *m_pMostRecentIbRow;
	IBROW              *m_pFreeIbRows;
	void               *m_pPagesToFree;
	BYTE               *m_pPagePos;
	BYTE               *m_pPageEnd;
	ULONG               m_cRows;
	ULONG               m_cColumns;
	ULONG               m_cbRowBuffer;
	isc_stmt_handle     m_iscStmtHandle;
	int                 m_iscStmtType;

	ULONG               m_ulAccessorSequence;
	bool                m_fUserCommand;
	bool                m_fEndOfRowset;
	bool                m_fPrefetchedRowset;
	bool                m_fPreserveOnCommit;
	bool                m_fIColumnsRowset;

	HRESULT BeginMethod(
		REFIID   riid);

	HRESULT FetchNextIbRows(
		HCHAPTER   hChapter, 
		LONG       lRowsOffset, 
		LONG       cRows, 
		ULONG     *pcRowsObtained, 
		PIBROW    *prghRows);

	HRESULT WINAPI InternalQueryInterface(
		CInterBaseRowset         *pThis,
		const _ATL_INTMAP_ENTRY  *pEntries, 
		REFIID                    iid, 
		void                    **ppvObject);
};

#endif /* _IBOLEDB_ROWSET_H_ */

