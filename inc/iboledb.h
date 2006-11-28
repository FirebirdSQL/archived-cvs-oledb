/*
 * Firebird Open Source OLE DB driver
 *
 * IBOLEDB.H - Include file for OLE DB session object definition.
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

#ifndef _IBOLEDB_IBOLEDB_H_
#define _IBOLEDB_IBOLEDB_H_

enum PROPENUMIBASE
{	DBPROP_IBASE_SESS_WAIT = 0x0010,
	DBPROP_IBASE_SESS_RECORDVERSIONS = 0x0011,
	DBPROP_IBASE_SESS_READONLY = 0x0012
};

#ifdef IBINITCONSTANTS
extern const GUID DBPROPSET_IBASE_SESSION =
				{0xE5336D6A,0xEDF2,0x404F,{0x83,0xAD,0xC1,0xA6,0x94,0xE7,0x1F,0xF5}};
extern const CLSID CLSID_InterBaseProvider =
				{0xBF4B098D,0x5A38,0x11D2,{0xB2,0x60,0x44,0x45,0x53,0x54,0x00,0x01}};
extern const IID IID_InterBaseRowset =
				{0xBF4B098D,0x5A38,0x11D2,{0xB2,0x60,0x44,0x45,0x53,0x54,0x00,0x02}};
extern const char *PROGID_InterBaseProvider = "IbOleDb";
extern const char *PROGID_InterBaseProvider_Version = "IbOleDb.1";
#else  /* IBINITCONSTANTS */
extern const GUID DBPROPSET_IBASE_SESSION;
extern const CLSID CLSID_InterBaseProvider;
extern const IID IID_InterBaseRowset;
extern const char *PROGID_InterBaseProvider;
extern const char *PROGID_InterBaseProvider_Version;
#endif /* IBINITCONSTANTS */

#endif /* _IBOLEDB_IBOLEDB_H_ */
