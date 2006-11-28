/*
 * Firebird Open Source OLE DB driver
 *
 * COMMAND.H - Include file for OLE DB command object definition.
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

#ifndef _IBOLEDB_COMMAND_H_
#define _IBOLEDB_COMMAND_H_

enum enumCOMMANDSTATE{
	CommandStateUninitialized,
	CommandStateInitial,
	CommandStateUnprepared,
	CommandStatePrepared,
	CommandStateExecuted
};

class ATL_NO_VTABLE  CInterBaseCommand :
	public CComObjectRoot,
	public IAccessor,
	public IColumnsInfo,
	public ICommandPrepare,
	public ICommandProperties,
	public ICommandText,
	public ICommandWithParameters,
	public IConvertType,
#ifdef ICOLUMNSROWSET
	public IColumnsRowset,
#endif /* ICOLUMNSROWSET */
	public ISupportErrorInfo
{
public:
	CInterBaseCommand();
	~CInterBaseCommand();


/* Internal Methods */
	CRITICAL_SECTION *GetCriticalSection(void);

	HRESULT DeleteRowset(
		isc_stmt_handle);

	HRESULT DescribeBindings(void);

	HRESULT Initialize(
		CInterBaseSession *, 
		isc_db_handle *, 
		bool );

	DBSTATUS GetIbValue(
		DBBINDING	  *pDbBinding,
		BYTE		  *pData,
		XSQLVAR		  *pSqlVar,
		BYTE		  *pSqlData,
		short		   SqlInd);

	HRESULT MakeColumnMap(
		XSQLDA		  *Sqlda, 
		IBCOLUMNMAP	 **ppColumnMap, 
		ULONG		  *pcbBuffer);

/* The ATL interface map */
BEGIN_COM_MAP(CInterBaseCommand)
	COM_INTERFACE_ENTRY(IAccessor)
	COM_INTERFACE_ENTRY(IColumnsInfo)
#ifdef ICOLUMNSROWSET
	COM_INTERFACE_ENTRY(IColumnsRowset)
#endif /* ICOLUMNSROWSET */
	COM_INTERFACE_ENTRY(ICommandPrepare)
	COM_INTERFACE_ENTRY(ICommand)
	COM_INTERFACE_ENTRY(ICommandProperties)
	COM_INTERFACE_ENTRY(ICommandText)
	COM_INTERFACE_ENTRY(ICommandWithParameters)
	COM_INTERFACE_ENTRY(IConvertType)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

	DECLARE_GET_CONTROLLING_UNKNOWN()
	DECLARE_AGGREGATABLE(CInterBaseCommand)

/* IAccessor */
	virtual HRESULT STDMETHODCALLTYPE AddRefAccessor( 
		HACCESSOR   hAccessor,
		ULONG      *pcRefCount);
        
	virtual HRESULT STDMETHODCALLTYPE CreateAccessor( 
		DBACCESSORFLAGS   dwAccessorFlags,
		ULONG             cBindings,
		const DBBINDING   rgBindings[  ],
		ULONG             cbRowSize,
		HACCESSOR        *phAccessor,
		DBBINDSTATUS      rgStatus[  ]);
        
	virtual HRESULT STDMETHODCALLTYPE GetBindings( 
		HACCESSOR         hAccessor,
		DBACCESSORFLAGS  *pdwAccessorFlags,
		ULONG            *pcBindings,
		DBBINDING       **prgBindings);
        
	virtual HRESULT STDMETHODCALLTYPE ReleaseAccessor( 
		HACCESSOR   hAccessor,
		ULONG      *pcRefCount);

#ifdef ICOLUMNSROWSET
/* IColumnsRowset */
HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetAvailableColumns(
		ULONG  *pcOptColumns,
		DBID  **prgOptColumns);

HRESULT STDMETHODCALLTYPE CInterBaseCommand::GetColumnsRowset(
		IUnknown    *pUnkOuter,
		ULONG        cOptColumns,
		const DBID   rgOptColumns[],
		REFIID       riid,
		ULONG        cPropertySets,
		DBPROPSET    rgPropertySets[],
		IUnknown   **ppColRowset);
#endif /* ICOLUMNSROWSET */

/* ICommand */
	virtual HRESULT STDMETHODCALLTYPE Cancel(void);
        
	virtual HRESULT STDMETHODCALLTYPE Execute( 
		IUnknown  *pUnkOuter,
		REFIID     riid,
		DBPARAMS  *pParams,
		LONG      *pcRowsAffected,
		IUnknown **ppRowset);
        
	virtual HRESULT STDMETHODCALLTYPE GetDBSession( 
		REFIID     riid,
		IUnknown **ppSession);

/* ICommandPrepare */
	virtual HRESULT STDMETHODCALLTYPE Prepare(
		ULONG   cExpectedRuns);

	virtual HRESULT STDMETHODCALLTYPE Unprepare(void);

/* ICommandProperties */
	virtual HRESULT STDMETHODCALLTYPE GetProperties( 
		const ULONG         cPropertyIDSets,
		const DBPROPIDSET   rgPropertyIDSets[],
		ULONG              *pcPropertySets,
		DBPROPSET         **prgPropertySets);
        
	virtual HRESULT STDMETHODCALLTYPE SetProperties( 
		ULONG       cPropertySets,
		DBPROPSET   rgPropertySets[]);

/* ICommandText */
	virtual HRESULT STDMETHODCALLTYPE GetCommandText( 
		GUID      *pguidDialect,
		LPOLESTR  *ppwszCommand);
        
	virtual HRESULT STDMETHODCALLTYPE SetCommandText( 
		REFGUID     rguidDialect,
		LPCOLESTR   pwszCommand);

/* ICommandWithParameters */
	virtual HRESULT STDMETHODCALLTYPE GetParameterInfo(
		ULONG        *pcParams,
		DBPARAMINFO **prgParamInfo,
		OLECHAR     **ppNamesBuffer);

	virtual HRESULT STDMETHODCALLTYPE MapParameterNames(
		ULONG           cParamNames,
		const OLECHAR  *rgParamNames[],
		LONG            rgParamOrdinals[]);

	virtual HRESULT STDMETHODCALLTYPE SetParameterInfo(
		ULONG                  cParams,
		const ULONG            rgParamOrdinals[],
		const DBPARAMBINDINFO  rgParamBindInfo[]);

/* IColumnsInfo */
	virtual HRESULT STDMETHODCALLTYPE GetColumnInfo(
		ULONG         *pcColumns, 
		DBCOLUMNINFO **prgInfo, 
		OLECHAR      **ppStringsBuffer);
 
	virtual HRESULT STDMETHODCALLTYPE MapColumnIDs(
		ULONG       cColumnIDs, 
		const DBID  rgColumnIDs[], 
		ULONG       rgColumns[]);

/* IConvertType */
	virtual HRESULT STDMETHODCALLTYPE CanConvert( 
		DBTYPE          wFromType,
		DBTYPE          wToType,
		DBCONVERTFLAGS  dwConvertFlags);

/* ISupportErrorInfo */
	STDMETHODIMP InterfaceSupportsErrorInfo(
		REFIID iid);

	BEGIN_IBPROPERTY_MAP
	BEGIN_IBPROPERTY_SET(DBPROPSET_ROWSET)
		IBPROPERTY_INFO_ENTRY_RW(COMMANDTIMEOUT, (DWORD)0)
#ifdef ICOLUMNSROWSET
		IBPROPERTY_INFO_ENTRY_RW(IColumnsRowset, VARIANT_TRUE)
#else  /* ICOLUMNSROWSET */
		IBPROPERTY_INFO_ENTRY_RO(IColumnsRowset, VARIANT_FALSE)
#endif /* ICOLUMNSROWSET (else) */
		IBPROPERTY_INFO_ENTRY_RO(BLOCKINGSTORAGEOBJECTS, VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(BOOKMARKS, VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(CANFETCHBACKWARDS, VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(CANHOLDROWS, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(CANSCROLLBACKWARDS, VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IAccessor, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(IColumnsInfo, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(IConvertType, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(IMultipleResults,  VARIANT_FALSE)
//		IBPROPERTY_INFO_ENTRY_RO(IRow,              VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowset,           VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetChange,     VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetIdentity,   VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetInfo,       VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetLocate,     VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetRefresh,    VARIANT_FALSE)
//		IBPROPERTY_INFO_ENTRY_RO(IRowsetResync,     VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(IRowsetUpdate,     VARIANT_FALSE)
		IBPROPERTY_INFO_ENTRY_RO(ISequentialStream, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(ISupportErrorInfo, VARIANT_TRUE)
		IBPROPERTY_INFO_ENTRY_RO(MAXOPENROWS, (DWORD)0)
		IBPROPERTY_INFO_ENTRY_RO(MAXROWS, (DWORD)0)
		IBPROPERTY_INFO_ENTRY_RO(MEMORYUSAGE, (DWORD)0)
		IBPROPERTY_INFO_ENTRY_RO(ROWTHREADMODEL, (DWORD)(DBPROPVAL_RT_FREETHREAD | DBPROPVAL_RT_SINGLETHREAD) )
		IBPROPERTY_INFO_ENTRY_RO(UPDATABILITY, (DWORD)0)
	END_IBPROPERTY_SET(DBPROPSET_ROWSET)
END_IBPROPERTY_MAP

private:
	CInterBaseError    *m_pInterBaseError;
	enumCOMMANDSTATE    m_eCommandState;
	bool                m_fUserCommand;
	bool                m_fExecuting;
	bool                m_fHaveKeyColumns;
	ULONG               m_cOpenRowsets;
	LPSTR               m_pszCommand;
	CRITICAL_SECTION    m_cs;
	CInterBaseSession  *m_pOwningSession;
	ISC_STATUS          m_iscStatus[20];
	isc_db_handle      *m_piscDbHandle;
	isc_tr_handle      *m_piscTransHandle;
	isc_stmt_handle     m_iscStmtHandle;
	
	XSQLDA             *m_pOutSqlda;
	ULONG               m_cColumns;
	IBCOLUMNMAP        *m_pIbColumnMap;
	ULONG               m_cbRowBuffer;

	XSQLDA             *m_pInSqlda;
	ULONG               m_cParams;
	ULONG               m_cbParamBuffer;
	IBCOLUMNMAP        *m_pIbParamMap;
	IBACCESSOR         *m_pIbAccessorsToRelease;
	ULONG               m_cIbAccessorCount;
	int                 m_iscStmtType;

	HRESULT BeginMethod(
		REFIID   riid);

	HRESULT WINAPI InternalQueryInterface(
		CInterBaseCommand        *pThis,
		const _ATL_INTMAP_ENTRY  *pEntries, 
		REFIID                    iid, 
		void                    **ppvObject);
};
#endif /* _IBOLEDB_COMMAND_H_ (not) */

