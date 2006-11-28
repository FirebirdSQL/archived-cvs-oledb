/*
 * Firebird Open Source OLE DB driver
 *
 * PROPERTY.H - Include file for OLE DB property implementation
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

#ifndef _IBOLEDB_PROPERTY_H_
#define _IBOLEDB_PROPERTY_H_

typedef struct {
	DBPROPID        dwID;
	DBPROPOPTIONS   dwFlags;
	bool            fIsInitialized;
	bool            fIsSet;
	LPOLESTR        wszName;
	VARTYPE         vt;
	union {
		DWORD      dwVal;
		LPOLESTR   wszVal;
	};
	VARIANTARG      vtVal;
} IBPROPERTY;

typedef IBPROPERTY * (CALLBACK *IBPROPCALLBACK)(REFIID, ULONG *);

HRESULT GetIbProperties(
	const ULONG         cPropertyIDSets,
	const DBPROPIDSET   rgPropertyIDSets[],
	ULONG              *pcPropertySets,
	DBPROPSET         **prgPropertySets,
	IBPROPCALLBACK      pfnGetIbProperty);

HRESULT SetIbProperties( 
	ULONG            cPropertySets,
	DBPROPSET        rgPropertySets[],
	IBPROPCALLBACK   pfnGetIbProperty);

HRESULT ClearIbProperties(
	IBPROPERTY  *pIbProperties,
	ULONG        cIbProperties);

VARIANTARG *GetIbPropertyValue(IBPROPERTY *pIbProperty);
DBPROPSTATUS SetIbPropertyValue(IBPROPERTY *pIbProperty, VARIANT *pvValue);

IBPROPERTY *LookupIbProperty(IBPROPERTY *, DWORD, DWORD);

#define MAX_PROPERTY_NAME_LENGTH		35

#define PROVIDER_PROPERTY_NAME(b) (PROVIDER_NAME L":" b)

#define DEFINE_DSO_PROPERTY(d,t,n) DEFINE_PROPERTY(d,t,n,DBPROPFLAGS_DATASOURCE)
#define DEFINE_DSOINFO_PROPERTY(d,t,n) DEFINE_PROPERTY(d,t,n,DBPROPFLAGS_DATASOURCEINFO)
#define DEFINE_DSOINIT_PROPERTY(d,t,n) DEFINE_PROPERTY(d,t,n,DBPROPFLAGS_DBINIT)
#define DEFINE_SESSION_PROPERTY(d,t,n) DEFINE_PROPERTY(d,t,n,DBPROPFLAGS_SESSION)
#define DEFINE_ROWSET_PROPERTY(d,t,n) DEFINE_PROPERTY(d,t,n,DBPROPFLAGS_ROWSET)

#ifdef DEFINE_IBOLEDB_PROPERTIES
#define DEFINE_PROPERTY(PropId,Type,Name,Flags) \
	extern const VARTYPE  PropId##_type = Type; \
	extern const LPOLESTR PropId##_name = (LPOLESTR)(Name); \
	extern const DBPROPFLAGS   PropId##_flags = Flags
#else  /* DEFINE_IBOLEDB_PROPERTIES */
#define DEFINE_PROPERTY(PropId,Type,Name,Flags) \
	extern const VARTYPE     PropId##_type; \
	extern const LPOLESTR    PropId##_name; \
	extern const DBPROPFLAGS PropId##_flags
#endif /* DEFINE_IBOLEDB_PROPERTIES (else) */

extern const OLEDBDECLSPEC GUID DBPROPSET_IBINIT = { 0xbef54d96, 0x84a5, 0x4a0c, { 0xaf, 0x1a, 0x2b, 0x3c, 0xb6, 0xe8, 0x51, 0xdc } };

/* Data Source Properties */
DEFINE_DSO_PROPERTY(CURRENTCATALOG, VT_BSTR, L"Current Catalog");

/* Data Source Information Properties */
DEFINE_DSOINFO_PROPERTY(ACTIVESESSIONS, VT_I4, L"Active Sessions");
DEFINE_DSOINFO_PROPERTY(BYREFACCESSORS, VT_BOOL, L"Pass By Ref Accessors");
DEFINE_DSOINFO_PROPERTY(CATALOGUSAGE, VT_I4, L"Catalog Usage");
DEFINE_DSOINFO_PROPERTY(COMSERVICES, VT_I4, L"COM Service Support");
DEFINE_DSOINFO_PROPERTY(CONCATNULLBEHAVIOR, VT_I4, L"NULL Concatenation Behavior");
DEFINE_DSOINFO_PROPERTY(DATASOURCEREADONLY, VT_BOOL, L"Read-Only Data Source");
DEFINE_DSOINFO_PROPERTY(DBMSNAME, VT_BSTR, L"DBMS Name");
DEFINE_DSOINFO_PROPERTY(DSOTHREADMODEL,              VT_I4, L"Data Source Object Threading Model");
DEFINE_DSOINFO_PROPERTY(IDENTIFIERCASE,              VT_I4, L"Identifier Case Sensitivity");
DEFINE_DSOINFO_PROPERTY(MULTIPLERESULTS,             VT_I4, L"Multiple Results");
DEFINE_DSOINFO_PROPERTY(MULTIPLESTORAGEOBJECTS,      VT_I4, L"Multiple Storage Objects");
DEFINE_DSOINFO_PROPERTY(NULLCOLLATION,               VT_I4, L"NULL Collation Order");
DEFINE_DSOINFO_PROPERTY(OLEOBJECTS,                  VT_I4, L"OLE Object Support");
DEFINE_DSOINFO_PROPERTY(OPENROWSETSUPPORT,           VT_I4, L"Open Rowset Support");
DEFINE_DSOINFO_PROPERTY(OUTPUTPARAMETERAVAILABILITY, VT_I4, L"Output Parameter Availability");
DEFINE_DSOINFO_PROPERTY(PROCEDURETERM, VT_BSTR, L"Procedure Term");
DEFINE_DSOINFO_PROPERTY(PROVIDERFRIENDLYNAME, VT_BSTR, L"Provider Friendly Name");
DEFINE_DSOINFO_PROPERTY(PROVIDERNAME,     VT_BSTR, L"Provider Name");
DEFINE_DSOINFO_PROPERTY(PROVIDEROLEDBVER, VT_BSTR, L"OLE DB Version");
DEFINE_DSOINFO_PROPERTY(PROVIDERVER,      VT_BSTR, L"Provider Version");
DEFINE_DSOINFO_PROPERTY(QUOTEDIDENTIFIERCASE, VT_I4, L"Quoted Identifier Sensitivity");
DEFINE_DSOINFO_PROPERTY(SQLSUPPORT, VT_I4, L"SQL Support");
DEFINE_DSOINFO_PROPERTY(STRUCTUREDSTORAGE, VT_I4, L"Structured Storage");
DEFINE_DSOINFO_PROPERTY(SUPPORTEDTXNISOLEVELS, VT_I4, L"Isolation Levels");
DEFINE_DSOINFO_PROPERTY(TABLETERM, VT_BSTR, L"Table Term");

/* Data Source Initialization Properties */
DEFINE_DSOINIT_PROPERTY(AUTH_PASSWORD, VT_BSTR, L"Password");
DEFINE_DSOINIT_PROPERTY(AUTH_USERID, VT_BSTR, L"User ID");
DEFINE_DSOINIT_PROPERTY(INIT_DATASOURCE, VT_BSTR, L"Data Source");
DEFINE_DSOINIT_PROPERTY(INIT_HWND, VT_I4, L"Window Handle");
DEFINE_DSOINIT_PROPERTY(INIT_LOCATION, VT_BSTR, L"Location");
DEFINE_DSOINIT_PROPERTY(AUTH_PERSIST_SENSITIVE_AUTHINFO, VT_BOOL, L"Persist Security Info");
DEFINE_DSOINIT_PROPERTY(INIT_PROMPT, VT_I2, L"Prompt");
DEFINE_DSOINIT_PROPERTY(INIT_PROVIDERSTRING, VT_BSTR, L"Extended Properties");
DEFINE_DSOINIT_PROPERTY(INIT_TIMEOUT, VT_I4, L"Connect Timeout");

/* Rowset Properties (Read Only) */
DEFINE_ROWSET_PROPERTY(ABORTPRESERVE, VT_BOOL, L"Preserve on Abort");
DEFINE_ROWSET_PROPERTY(BLOCKINGSTORAGEOBJECTS, VT_BOOL, L"Blocking Storage Objects");
DEFINE_ROWSET_PROPERTY(BOOKMARKS, VT_BOOL, L"Use Bookmarks");
DEFINE_ROWSET_PROPERTY(CANFETCHBACKWARDS, VT_BOOL, L"Fetch Backward");
DEFINE_ROWSET_PROPERTY(CANHOLDROWS, VT_BOOL, L"Hold Rows");
DEFINE_ROWSET_PROPERTY(CANSCROLLBACKWARDS, VT_BOOL, L"Scroll Backward");
DEFINE_ROWSET_PROPERTY(COMMANDTIMEOUT, VT_I4, L"Command Time Out");
DEFINE_ROWSET_PROPERTY(COMMITPRESERVE, VT_BOOL, L"Preserve on Commit");
DEFINE_ROWSET_PROPERTY(DEFERRED, VT_BOOL, L"Defer Column");
DEFINE_ROWSET_PROPERTY(IAccessor, VT_BOOL, L"IAccessor");
DEFINE_ROWSET_PROPERTY(IChapteredRowset, VT_BOOL, L"IChapteredRowset");
DEFINE_ROWSET_PROPERTY(IColumnsInfo, VT_BOOL, L"IColumnsInfo");
DEFINE_ROWSET_PROPERTY(IColumnsRowset, VT_BOOL, L"IColumnsRowset");
DEFINE_ROWSET_PROPERTY(IConnectionPointContainer, VT_BOOL, L"IConnectionPointContainer");
DEFINE_ROWSET_PROPERTY(IConvertType,    VT_BOOL, L"IConvertType");
DEFINE_ROWSET_PROPERTY(IRow,            VT_BOOL, L"IRow");
DEFINE_ROWSET_PROPERTY(IRowset,         VT_BOOL, L"IRowset");
DEFINE_ROWSET_PROPERTY(IRowsetChange,   VT_BOOL, L"IRowsetChange");
DEFINE_ROWSET_PROPERTY(IRowsetIdentity, VT_BOOL, L"IRowsetIdentity");
DEFINE_ROWSET_PROPERTY(IRowsetInfo,     VT_BOOL, L"IRowsetInfo");
DEFINE_ROWSET_PROPERTY(IRowsetLocate,   VT_BOOL, L"IRowsetLocate");
DEFINE_ROWSET_PROPERTY(IRowsetRefresh,  VT_BOOL, L"IRowsetRefresh");
DEFINE_ROWSET_PROPERTY(IRowsetResync,   VT_BOOL, L"IRowsetResync");
DEFINE_ROWSET_PROPERTY(IRowsetUpdate,   VT_BOOL, L"IRowsetUpdate");
DEFINE_ROWSET_PROPERTY(IMultipleResults, VT_BOOL, L"IMultipleResults");
DEFINE_ROWSET_PROPERTY(ISequentialStream, VT_BOOL, L"ISequentialStream");
DEFINE_ROWSET_PROPERTY(ISupportErrorInfo, VT_BOOL, L"ISupportErrorInfo");
DEFINE_ROWSET_PROPERTY(MAXOPENROWS, VT_I4, L"Maximum Open Rows");
DEFINE_ROWSET_PROPERTY(MAXROWS, VT_I4, L"Maximum Rows");
DEFINE_ROWSET_PROPERTY(MEMORYUSAGE, VT_I4, L"Memory Usage");
DEFINE_ROWSET_PROPERTY(ROWTHREADMODEL, VT_I4, L"Row Threading Model");	//DBPROPVAL_RT_FREETHREAD
DEFINE_ROWSET_PROPERTY(QUICKRESTART, VT_BOOL, L"Quick Restart");	//False
DEFINE_ROWSET_PROPERTY(UPDATABILITY, VT_I4, L"Updatability");		//0

/* Session Properties */
DEFINE_SESSION_PROPERTY(SESS_AUTOCOMMITISOLEVELS, VT_I4, L"Autocommit Isolation Levels");
DEFINE_SESSION_PROPERTY(IBASE_SESS_WAIT, VT_BOOL, PROVIDER_PROPERTY_NAME(L"Wait"));
DEFINE_SESSION_PROPERTY(IBASE_SESS_RECORDVERSIONS, VT_BOOL, PROVIDER_PROPERTY_NAME(L"Record Versions"));
DEFINE_SESSION_PROPERTY(IBASE_SESS_READONLY, VT_BOOL, PROVIDER_PROPERTY_NAME(L"Read Only"));

/* Property definition macros */
#define BEGIN_IBPROPERTY_MAP static IBPROPERTY * CALLBACK _getIbProperties(REFIID iid, ULONG *pcRows) {
#define BEGIN_IBPROPERTY_SET(guid) if (IsEqualGUID(iid, guid)) { \
	static IBPROPERTY PropSet##guid [] = {

#define IBPROPERTY_INFO_ENTRY_RO(dwPropID, value) DBPROP_##dwPropID, DBPROPFLAGS_READ | dwPropID##_flags, false, false, dwPropID##_name, dwPropID##_type, (DWORD)value, {0,0,0,0},
#define IBPROPERTY_INFO_ENTRY_RW(dwPropID, value) DBPROP_##dwPropID, DBPROPFLAGS_READ | DBPROPFLAGS_WRITE | dwPropID##_flags, false, false, dwPropID##_name, dwPropID##_type, (DWORD)value, {0,0,0,0},

#define END_IBPROPERTY_SET(guid) }; \
		if(pcRows != NULL) \
			*pcRows = sizeof(PropSet##guid)/sizeof(IBPROPERTY); \
		return PropSet##guid; \
	}

#define END_IBPROPERTY_MAP *pcRows = 0; return NULL;	}

#define PROPERTY_IS_SET(p,i)

#endif /* _IBOLEDB_PROPERTY_H_ (not) */