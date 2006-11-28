/*
 * Firebird Open Source OLE DB driver
 *
 * IBOLEDB.CPP - Source file for various things required by the DLL.
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

#include "../inc/common.h"
#include "../inc/resource.h"
#define IBINITCONSTANTS
#include "../inc/iboledb.h"
#include "../inc/property.h"
#include "../inc/ibaccessor.h"
#include "../inc/command.h"
#include "../inc/session.h"
#include "../inc/dso.h"

HINSTANCE g_hInstance = NULL;
CComModule _Module;

/* ATL implementation code */
#include <atlimpl.cpp>

BEGIN_OBJECT_MAP(ObjectMap)
	OBJECT_ENTRY(CLSID_InterBaseProvider, CInterBaseDSO)
END_OBJECT_MAP()

static BOOL WINAPI DllMain2(HANDLE hInst, ULONG ulReason, LPVOID lpReserved);

static HRESULT DllRegisterOLEDBProvider(void)
{
	static const TCHAR szOLEDBProvider[] = _T("OLE DB Provider");
	static const TCHAR szOLEDBServices[] = _T("OLEDB_Services");
	
	CRegKey keyCLSID;
	LONG 	lRes;
	HRESULT	hRes;

	// Open HKEY_CLASSES_ROOT\CLSID
	if ((lRes = keyCLSID.Open(HKEY_CLASSES_ROOT, _T("CLSID"))) == ERROR_SUCCESS)
	{
		// Open the entry for our particular class ID.
		CRegKey 	keyProgID;
		LPOLESTR 	lpOleStrProgID;
		ULONG 		cch;
		LPSTR		lpszProgID;

		StringFromCLSID(CLSID_InterBaseProvider, &lpOleStrProgID);
		cch = WideCharToMultiByte(CP_ACP, 0, lpOleStrProgID, -1, 0, 0, 0, 0);

		_ASSERT(cch != 0);

		if ((lpszProgID = (LPSTR)alloca(cch + 1)) == NULL)
		{
			CoTaskMemFree(lpOleStrProgID);
			return E_OUTOFMEMORY;
		}
		
		if (WideCharToMultiByte(CP_ACP, 0, lpOleStrProgID, -1, lpszProgID, cch, 0, 0) == 0)
		{
			CoTaskMemFree(lpOleStrProgID);
			return E_FAIL;
		}

		CoTaskMemFree(lpOleStrProgID);

		if ((lRes = keyProgID.Open(keyCLSID, lpszProgID)) == ERROR_SUCCESS)
		{
			// Create the "OLE DB Provider" key.
			CRegKey	keyProvider;

			keyProgID.SetValue((DWORD)0xfffffffb, szOLEDBServices);

			if ((lRes = keyProvider.Create(keyProgID, szOLEDBProvider)) == ERROR_SUCCESS)
				lRes = keyProvider.SetValue(_T(PROGID_InterBaseProvider));
		}
	}

	hRes = (lRes == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lRes);

	return hRes;
}

/* DLL Entry Point */

extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	if (dwReason == DLL_PROCESS_ATTACH)
	{
		_Module.Init(ObjectMap, hInstance);
		DisableThreadLibraryCalls(hInstance);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
		_Module.Term();

	return DllMain2(hInstance, dwReason, lpReserved);
}

STDAPI DllCanUnloadNow(void)
{
	
	if (_Module.GetLockCount()==0) 
	{
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.GetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	HRESULT hRes = S_OK;
	hRes = _Module.RegisterServer(FALSE);

	if (!FAILED(hRes))
		hRes = DllRegisterOLEDBProvider();

	return hRes;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT hRes = S_OK;

	_Module.UnregisterServer();

	return hRes;
}

static BOOL IbOleDbInit(void)
{
	
	BOOL fOk = TRUE;

	return fOk;
}

static void IbOleDbShutdown(void)
{
}

static BOOL WINAPI DllMain2(HANDLE hInst, ULONG ulReason, LPVOID lpReserved)
{
	BOOL fOk = TRUE;

	switch (ulReason)
	{
	 case DLL_PROCESS_ATTACH:
	 	_ASSERT(!g_hInstance);
	 	g_hInstance = (HINSTANCE)hInst;

	 	if (fOk = IbOleDbInit())
	 		DbgTrace("IbOleDb Initialization Succeeded.\n");
	 	else
	 		DbgTrace("IbOleDb Initialization Failed!!!\n");
	 	break;
	 	
	 case DLL_PROCESS_DETACH:
	 	_ASSERT(g_hInstance);

	 	DbgTrace("IbOleDb Terminating...\n");

	 	IbOleDbShutdown();
	 	break;
	}

	return fOk;
}


