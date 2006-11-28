/*
 * Firebird Open Source OLE DB driver
 *
 * SESSION.H - Include file for OLE DB session object definition
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

#ifndef _IBOLEDB_SESSION_H_
#define _IBOLEDB_SESSION_H_

class CInterBaseSession:
public CComObjectRoot,
       public IOpenRowset,
       public IGetDataSource,
       public ISessionProperties,
       public IDBCreateCommand,
#ifdef IDBSCHEMAROWSET
       public IDBSchemaRowset,
#endif /* IDBSCHEMAROWSET */
       public ISupportErrorInfo,
       public ITransactionLocal
	
{
public:
	CInterBaseSession();
	~CInterBaseSession();

	void *AllocPage( ULONG ulPageSize );
	void FreePage( void *pPage );
	HRESULT Initialize(CInterBaseDSO *, isc_db_handle *);
	HRESULT AutoCommit( bool fSuccess );
	HRESULT GetTransHandle( isc_tr_handle ** );
	HRESULT ReleaseTransHandle( isc_tr_handle ** );
	HRESULT InternalCommit( bool, DWORD );
	HRESULT InternalAbort( bool );

	/* The ATL interface map */
	BEGIN_COM_MAP(CInterBaseSession)
		COM_INTERFACE_ENTRY(IOpenRowset)
		COM_INTERFACE_ENTRY(IGetDataSource)
		COM_INTERFACE_ENTRY(ISessionProperties)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(IDBCreateCommand)
#ifdef IDBSCHEMAROWSET
		COM_INTERFACE_ENTRY(IDBSchemaRowset)
#endif /* IDBSCHEMAROWSET */
		COM_INTERFACE_ENTRY(ITransaction)
		COM_INTERFACE_ENTRY(ITransactionLocal)
	END_COM_MAP()

	DECLARE_GET_CONTROLLING_UNKNOWN()
	DECLARE_POLY_AGGREGATABLE(CInterBaseSession)

/* IGetDataSource */
	virtual HRESULT STDMETHODCALLTYPE GetDataSource(
		REFIID     riid,
		IUnknown **ppDataSource);

/* IOpenRowset */
	virtual HRESULT STDMETHODCALLTYPE OpenRowset(
		IUnknown  *pUnkOuter,
		DBID       *pTableID,
		DBID       *pIndexID,
		REFIID      riid,
		ULONG       cProperties,
		DBPROPSET   rgProperties[  ],
		IUnknown  **ppRowset);

/* ISessionProperties */
	virtual HRESULT STDMETHODCALLTYPE GetProperties(
		ULONG               cPropertyIDSets,
		const DBPROPIDSET   rgPropertyIDSets[],
		ULONG              *pcPropertySets,
		DBPROPSET         **prgPropertySets);

	virtual HRESULT STDMETHODCALLTYPE SetProperties(
		ULONG       cPropertySets,
		DBPROPSET   rgPropertySets[]);

/* IDBCreateCommand */
	virtual HRESULT STDMETHODCALLTYPE CreateCommand(
		IUnknown  *pUnkOuter,
		REFIID     riid,
		IUnknown **ppCommand);

/* IDBSchemaRowset */
	virtual HRESULT STDMETHODCALLTYPE GetRowset (
		IUnknown       *pUnkOuter,
		REFGUID         rguidSchema,
		ULONG           cRestrictions,
		const VARIANT   rgRestrictions[],
		REFIID          riid,
		ULONG           cPropertySets,
		DBPROPSET       rgPropertySets[],
		IUnknown      **ppRowset);

	virtual HRESULT STDMETHODCALLTYPE GetSchemas (
		ULONG  *pcSchemas,
		GUID  **prgSchemas,
		ULONG **prgRestrictionSupport);

/* ISupportErrorInfo */
	STDMETHODIMP InterfaceSupportsErrorInfo(
		REFIID   iid);

/* ITransaction */
	virtual HRESULT STDMETHODCALLTYPE Abort(
		BOID  *pboidReason,
		BOOL   fRetaining,
		BOOL   fAsynch);

	virtual HRESULT STDMETHODCALLTYPE Commit(
		BOOL    fRetaining,
		DWORD   grfTC,
		DWORD   grfRM);

	virtual HRESULT STDMETHODCALLTYPE GetTransactionInfo(
		XACTTRANSINFO  *pInfo);

/* ITransactionLocal */
	virtual HRESULT STDMETHODCALLTYPE GetOptionsObject(
		ITransactionOptions **ppOptions);
		
	virtual HRESULT STDMETHODCALLTYPE StartTransaction(
		ISOLEVEL              isoLevel,
		ULONG                 isoFlags,
		ITransactionOptions  *pOtherOptions,
		ULONG                *pulTransactionLevel);

#define IBASE_SESS_WAIT_INDEX           0
#define IBASE_SESS_RECORDVERSIONS_INDEX 1
#define IBASE_SESS_READONLY_INDEX       2
BEGIN_IBPROPERTY_MAP
	BEGIN_IBPROPERTY_SET(DBPROPSET_SESSION)
		IBPROPERTY_INFO_ENTRY_RW(SESS_AUTOCOMMITISOLEVELS, (DWORD)(DBPROPVAL_TI_REPEATABLEREAD))
	END_IBPROPERTY_SET(DBPROPSET_SESSION)
	BEGIN_IBPROPERTY_SET(DBPROPSET_IBASE_SESSION)
		IBPROPERTY_INFO_ENTRY_RW(IBASE_SESS_WAIT, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RW(IBASE_SESS_RECORDVERSIONS, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RW(IBASE_SESS_READONLY, VARIANT_FALSE)
	END_IBPROPERTY_SET(DBPROPSET_IBASE_SESSION)
END_IBPROPERTY_MAP

private:
	CInterBaseError       *m_pInterBaseError;
	CInterBaseDSO         *m_pOwningDSO;
	isc_db_handle         *m_piscDbHandle;
	void                  *m_pFreePages;
	CRITICAL_SECTION       m_cs;
	ULONG                  m_cTransLocks;
	ULONG	               m_ulTransLevel;
	ISC_STATUS             m_iscStatus[20];
	isc_tr_handle          m_iscTransHandle;
	ISOLEVEL               m_isoLevel;
	bool                   m_fAutoCommit;

	HRESULT BeginMethod(
		REFIID   riid);

	HRESULT WINAPI InternalQueryInterface(
		CInterBaseSession        *pThis,
		const _ATL_INTMAP_ENTRY  *pEntries, 
		REFIID                    iid, 
		void                    **ppvObject);
};

#endif	/* _IBOLEDB_SESSION_H_ */
