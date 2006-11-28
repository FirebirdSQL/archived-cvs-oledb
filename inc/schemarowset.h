/*
 * Firebird Open Source OLE DB driver
 *
 * SCHEMAROWSET.H - Include file defines schema rowset columns and
 *                  a utility class used to fetch data from the 
 *                  InterBase system catalogs.
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

#ifndef _IBOLEDB_SCHEMAROWSET_H_
#define _IBOLEDB_SCHEMAROWSET_H_

#define PREFETCHED_STATUS(c, p) (short *)((BYTE *)(p) + (c).obStatus)
#define PREFETCHED_DATA(c, p) (void *)((BYTE *)(p) + (c).obData)

#include <string>
using namespace std;
typedef basic_string<char, char_traits<char>, allocator<char> > CIbString;
typedef HRESULT (*PREFETCHEDROWSETCALLBACK)( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );

HRESULT SchemaColumnsRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
HRESULT SchemaTablesRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
HRESULT SchemaProviderTypesRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );

#ifdef OPTIONAL_SCHEMA_ROWSETS
HRESULT SchemaForeignKeysRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
HRESULT SchemaIndexesRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
HRESULT SchemaPrimaryKeysRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
HRESULT SchemaProceduresRowset( isc_db_handle *piscDbHandle, CComPolyObject<CInterBaseRowset> *pRowset, ULONG cRestrictions, const VARIANT rgRestrictions[] );
#endif /* OPTIONAL_SCHEMA_ROWSETS */

extern const OLEDBDECLSPEC DBID IBCOLUMN_AUTOUNIQUEVALUE                   = {DB_NULLGUID, DBKIND_NAME, L"AUTO_UNIQUE_VALUE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_AUTOUPDATE                        = {DB_NULLGUID, DBKIND_NAME, L"AUTO_UPDATE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_BESTMATCH                         = {DB_NULLGUID, DBKIND_NAME, L"BEST_MATCH"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CARDINALITY                       = {DB_NULLGUID, DBKIND_NAME, L"CARDINALITY"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CASESENSITIVE                     = {DB_NULLGUID, DBKIND_NAME, L"CASE_SENSITIVE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CHARACTERMAXIMUMLENGTH            = {DB_NULLGUID, DBKIND_NAME, L"CHARACTER_MAXIMUM_LENGTH"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CHARACTEROCTETLENGTH              = {DB_NULLGUID, DBKIND_NAME, L"CHARACTER_OCTET_LENGTH"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CHARACTERSETCATALOG               = {DB_NULLGUID, DBKIND_NAME, L"CHARACTER_SET_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CHARACTERSETNAME                  = {DB_NULLGUID, DBKIND_NAME, L"CHARACTER_SET_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CHARACTERSETSCHEMA                = {DB_NULLGUID, DBKIND_NAME, L"CHARACTER_SET_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CLUSTERED                         = {DB_NULLGUID, DBKIND_NAME, L"CLUSTERED"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLLATION                         = {DB_NULLGUID, DBKIND_NAME, L"COLLATION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLLATIONCATALOG                  = {DB_NULLGUID, DBKIND_NAME, L"COLLATION_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLLATIONNAME                     = {DB_NULLGUID, DBKIND_NAME, L"COLLATION_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLLATIONSCHEMA                   = {DB_NULLGUID, DBKIND_NAME, L"COLLATION_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNDEFAULT                     = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_DEFAULT"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNFLAGS                       = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_FLAGS"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNGUID                        = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNHASDEFAULT                  = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_HASDEFAULT"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNNAME                        = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNPROPID                      = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_PROPID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_COLUMNSIZE                        = {DB_NULLGUID, DBKIND_NAME, L"COLUMN_SIZE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_CREATEPARAMS                      = {DB_NULLGUID, DBKIND_NAME, L"CREATE_PARAMS"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DATATYPE                          = {DB_NULLGUID, DBKIND_NAME, L"DATA_TYPE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DATECREATED                       = {DB_NULLGUID, DBKIND_NAME, L"DATE_CREATED"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DATEMODIFIED                      = {DB_NULLGUID, DBKIND_NAME, L"DATE_MODIFIED"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DATETIMEPRECISION                 = {DB_NULLGUID, DBKIND_NAME, L"DATETIME_PRECISION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DEFERRABILITY                     = {DB_NULLGUID, DBKIND_NAME, L"DEFERRABILITY"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DELETERULE                        = {DB_NULLGUID, DBKIND_NAME, L"DELETE_RULE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DESCRIPTION                       = {DB_NULLGUID, DBKIND_NAME, L"DESCRIPTION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DOMAINCATALOG                     = {DB_NULLGUID, DBKIND_NAME, L"DOMAIN_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DOMAINNAME                        = {DB_NULLGUID, DBKIND_NAME, L"DOMAIN_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_DOMAINSCHEMA                      = {DB_NULLGUID, DBKIND_NAME, L"DOMAIN_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FILLFACTOR                        = {DB_NULLGUID, DBKIND_NAME, L"FILL_FACTOR"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FILTERCONDITION                   = {DB_NULLGUID, DBKIND_NAME, L"FILTER_CONDITION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FIXEDPRECSCALE                    = {DB_NULLGUID, DBKIND_NAME, L"FIXED_PREC_SCALE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKCOLUMNGUID                      = {DB_NULLGUID, DBKIND_NAME, L"FK_COLUMN_GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKCOLUMNNAME                      = {DB_NULLGUID, DBKIND_NAME, L"FK_COLUMN_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKCOLUMNPROPID                    = {DB_NULLGUID, DBKIND_NAME, L"FK_COLUMN_PROPID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKNAME                            = {DB_NULLGUID, DBKIND_NAME, L"FK_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKTABLECATALOG                    = {DB_NULLGUID, DBKIND_NAME, L"FK_TABLE_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKTABLENAME                       = {DB_NULLGUID, DBKIND_NAME, L"FK_TABLE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_FKTABLESCHEMA                     = {DB_NULLGUID, DBKIND_NAME, L"FK_TABLE_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_GUID                              = {DB_NULLGUID, DBKIND_NAME, L"GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_INDEXCATALOG                      = {DB_NULLGUID, DBKIND_NAME, L"INDEX_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_INDEXNAME                         = {DB_NULLGUID, DBKIND_NAME, L"INDEX_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_INDEXSCHEMA                       = {DB_NULLGUID, DBKIND_NAME, L"INDEX_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_INITIALSIZE                       = {DB_NULLGUID, DBKIND_NAME, L"INITIAL_SIZE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_INTEGRATED                        = {DB_NULLGUID, DBKIND_NAME, L"INTEGRATED"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_ISFIXEDLENGTH                     = {DB_NULLGUID, DBKIND_NAME, L"IS_FIXEDLENGTH"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_ISLONG                            = {DB_NULLGUID, DBKIND_NAME, L"IS_LONG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_ISNULLABLE                        = {DB_NULLGUID, DBKIND_NAME, L"IS_NULLABLE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_LOCALTYPENAME                     = {DB_NULLGUID, DBKIND_NAME, L"LOCAL_TYPE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_LITERALPREFIX                     = {DB_NULLGUID, DBKIND_NAME, L"LITERAL_PREFIX"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_LITERALSUFFIX                     = {DB_NULLGUID, DBKIND_NAME, L"LITERAL_SUFFIX"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_MAXIMUMSCALE                      = {DB_NULLGUID, DBKIND_NAME, L"MAXIMUM_SCALE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_MINIMUMSCALE                      = {DB_NULLGUID, DBKIND_NAME, L"MINIMUM_SCALE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_NULLCOLLATION                     = {DB_NULLGUID, DBKIND_NAME, L"NULL_COLLATION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_NULLS                             = {DB_NULLGUID, DBKIND_NAME, L"NULLS"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_NUMERICPRECISION                  = {DB_NULLGUID, DBKIND_NAME, L"NUMERIC_PRECISION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_NUMERICSCALE                      = {DB_NULLGUID, DBKIND_NAME, L"NUMERIC_SCALE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_ORDINAL                           = {DB_NULLGUID, DBKIND_NAME, L"ORDINAL"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_ORDINALPOSITION                   = {DB_NULLGUID, DBKIND_NAME, L"ORDINAL_POSITION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PAGES                             = {DB_NULLGUID, DBKIND_NAME, L"PAGES"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKCOLUMNGUID                      = {DB_NULLGUID, DBKIND_NAME, L"PK_COLUMN_GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKCOLUMNNAME                      = {DB_NULLGUID, DBKIND_NAME, L"PK_COLUMN_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKCOLUMNPROPID                    = {DB_NULLGUID, DBKIND_NAME, L"PK_COLUMN_PROPID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKNAME                            = {DB_NULLGUID, DBKIND_NAME, L"PK_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKTABLECATALOG                    = {DB_NULLGUID, DBKIND_NAME, L"PK_TABLE_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKTABLENAME                       = {DB_NULLGUID, DBKIND_NAME, L"PK_TABLE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PKTABLESCHEMA                     = {DB_NULLGUID, DBKIND_NAME, L"PK_TABLE_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PRIMARYKEY                        = {DB_NULLGUID, DBKIND_NAME, L"PRIMARY_KEY"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PROCEDURECATALOG                  = {DB_NULLGUID, DBKIND_NAME, L"PROCEDURE_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PROCEDUREDEFINITION               = {DB_NULLGUID, DBKIND_NAME, L"PROCEDURE_DEFINITION"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PROCEDURENAME                     = {DB_NULLGUID, DBKIND_NAME, L"PROCEDURE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PROCEDURESCHEMA                   = {DB_NULLGUID, DBKIND_NAME, L"PROCEDURE_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_PROCEDURETYPE                     = {DB_NULLGUID, DBKIND_NAME, L"PROCEDURE_TYPE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_SEARCHABLE                        = {DB_NULLGUID, DBKIND_NAME, L"SEARCHABLE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_SORTBOOKMARKS                     = {DB_NULLGUID, DBKIND_NAME, L"SORT_BOOKMARKS"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLECATALOG                      = {DB_NULLGUID, DBKIND_NAME, L"TABLE_CATALOG"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLEGUID                         = {DB_NULLGUID, DBKIND_NAME, L"TABLE_GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLENAME                         = {DB_NULLGUID, DBKIND_NAME, L"TABLE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLEPROPID                       = {DB_NULLGUID, DBKIND_NAME, L"TABLE_PROPID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLESCHEMA                       = {DB_NULLGUID, DBKIND_NAME, L"TABLE_SCHEMA"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TABLETYPE                         = {DB_NULLGUID, DBKIND_NAME, L"TABLE_TYPE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TYPE                              = {DB_NULLGUID, DBKIND_NAME, L"TYPE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TYPEGUID                          = {DB_NULLGUID, DBKIND_NAME, L"TYPE_GUID"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TYPELIB                           = {DB_NULLGUID, DBKIND_NAME, L"TYPELIB"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_TYPENAME                          = {DB_NULLGUID, DBKIND_NAME, L"TYPE_NAME"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_UNIQUE                            = {DB_NULLGUID, DBKIND_NAME, L"UNIQUE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_UNSIGNEDATTRIBUTE                 = {DB_NULLGUID, DBKIND_NAME, L"UNSIGNED_ATTRIBUTE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_UPDATERULE                        = {DB_NULLGUID, DBKIND_NAME, L"UPDATE_RULE"};
extern const OLEDBDECLSPEC DBID IBCOLUMN_VERSION                           = {DB_NULLGUID, DBKIND_NAME, L"VERSION"};

typedef struct tagIBPREFETCHEDCOLUMNINFO
{
	DBCOLUMNFLAGS dwFlags;	//Flags
	short sqllen;			//Column Size
	short sqltype;			//Internal Type
	DBTYPE wType;			//External Type
	DBID columnid;			//DBID
	WCHAR *pwszColumnName;
} PrefetchedColumnInfo;

#define PREFECTHEDCOLUMNINFOCOUNT( v ) (sizeof( v ) / sizeof( PrefetchedColumnInfo ) )
#define EXPAND_DBID( dbid ) { EXPANDGUID(dbid.uGuid.guid), (DWORD)dbid.eKind, (LPOLESTR) dbid.uName.pwszName }
#define SIMPLE_DBID( n ) { DB_NULLGUID, 0, (LPOLESTR)(n) }

#define BEGIN_PREFETCHEDCOLUMNINFO( v ) \
static PrefetchedColumnInfo v[] = { 

#define PREFETCHEDCOLUMNINFO_ENTRY( flags, sqllen, dbtype, sqltype, dbid ) \
	PREFETCHEDCOLUMNINFO_ENTRY2( flags, sqllen, dbtype, sqltype, dbid, dbid.uName.pwszName )

#define PREFETCHEDCOLUMNINFO_ENTRY2( flags, sqllen, dbtype, sqltype, dbid, name ) \
{ (DBCOLUMNFLAGS)flags, sqllen, sqltype, (DBTYPE)dbtype, { EXPANDGUID(dbid.uGuid.guid), (DWORD)dbid.eKind, (LPOLESTR) dbid.uName.pwszName}, (LPOLESTR)name },

#define END_PREFETCHEDCOLUMNINFO \
};

enum BindTypeEnum
{
	BindTypeColumn,
	BindTypeParameter
};

class CSchemaRowsetHelper
{
public:
	CSchemaRowsetHelper( void );
	~CSchemaRowsetHelper( void );
	bool Init( isc_db_handle *iscDbHandle );
	bool Exec( void );
	bool Fetch( void );
	bool Bind( BindTypeEnum eBindType, short sColumn, short sqltype, short *psStatus, void *pData, short sLength );
	bool BindCol( short sColumn, short sqltype, short *psStatus, void *pData, short sLength );
	bool BindParam( short sParam, short sqltype, short *psStatus, void *pData, short sLength );
	bool Append( const char *, int );
	bool Append( const char * );
	bool AppendWhere( const char * );
	bool AppendWhereParam( const char * );
	void SetCommand( const char * );
	void SetColumnCount( short );
	const char *GetCommand( void );
private:
	CIbString         m_strCommand;
	isc_db_handle    *m_piscDbHandle;
	isc_tr_handle     m_iscTransHandle;
	isc_stmt_handle   m_iscStmtHandle;
	ISC_STATUS        m_iscStatus[20];
	int               m_rc;
	int               m_cPredicates;
	XSQLDA           *m_pInSqlda;
	XSQLDA           *m_pOutSqlda;
	short             m_cMaxColumns;
	bool              m_fNeedsWhere;
};

#endif /* _IBOLEDB_SCHEMAROWSET_H_ */