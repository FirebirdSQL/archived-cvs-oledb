/*
 * Firebird Open Source OLE DB driver
 *
 * IBERROR.CPP - Source file for error object methods.
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
#include "../inc/iboledb.h"
#include "../inc/util.h"
#include <msdaguid.h>

CInterBaseError::CInterBaseError() :
	m_pwszDesc(NULL),
	m_pwszDescEnd(NULL),
	m_cbDesc(0),
	m_riid(IID_NULL),
	m_fErrorPosted(false)
{
}

CInterBaseError::~CInterBaseError()
{
	if (m_pwszDesc)
		delete[] m_pwszDesc;
}

void CInterBaseError::ClearError(REFIID riid)
{
	SetErrorInfo(0, NULL);
	m_fErrorPosted = false;

	memcpy((void *)&m_riid, (void *)&riid, sizeof(GUID));
	
	if (m_pwszDescEnd = m_pwszDesc)
		*m_pwszDescEnd = L'\0';
}

HRESULT CInterBaseError::PostError(WCHAR *pwszError, DWORD dwMinor, HRESULT hr)
{
	ICreateErrorInfo	  *pCreateErrorInfo = NULL;
	IErrorInfo			  *pErrorInfo = NULL;
	IErrorRecords		  *pErrorRecords = NULL;
	WCHAR				  *pwszPROGID = NULL;
	ULONG				   cbPROGID;

	if (SUCCEEDED(CreateErrorInfo(&pCreateErrorInfo)))
	{
		ERRORINFO	errorinfo;
		errorinfo.clsid = CLSID_InterBaseProvider;
		errorinfo.dispid = NULL;
		errorinfo.dwMinor = dwMinor;
		// errorinfo.dwMinor = piscStatus[1];
		errorinfo.hrError = hr;
		errorinfo.iid = m_riid;
		pCreateErrorInfo->SetGUID(errorinfo.iid);

		cbPROGID = lstrlenA(PROGID_InterBaseProvider)+1;
		pwszPROGID = new WCHAR [mbstowcs(NULL,
										 PROGID_InterBaseProvider,
										 cbPROGID)+1];
		mbstowcs(pwszPROGID, PROGID_InterBaseProvider, cbPROGID);
		pCreateErrorInfo->SetSource(pwszPROGID);
		pCreateErrorInfo->SetDescription(pwszError);

		GetErrorInfo(0, &pErrorInfo);
		if (pErrorInfo == NULL)
		{
			if (FAILED(CoCreateInstance(
					CLSID_EXTENDEDERRORINFO,
					NULL,
					CLSCTX_INPROC_SERVER,
					IID_IErrorInfo,
					(LPVOID *) &pErrorInfo)))
				goto cleanup;
		}
		
		if (SUCCEEDED(pErrorInfo->QueryInterface(
				IID_IErrorRecords,
				(void **) &pErrorRecords)))
		{
			pErrorRecords->AddErrorRecord(
					&errorinfo,
					0,
					NULL,
					pCreateErrorInfo,
					0);
		}

	}
	cleanup:
	if (pErrorInfo)
	{
		SetErrorInfo(0, pErrorInfo);
		pErrorInfo->Release();
	}
	if (pCreateErrorInfo)
		pCreateErrorInfo->Release();
	if (pErrorRecords)
		pErrorRecords->Release();

	return hr;
}

HRESULT CInterBaseError::PostError(ISC_STATUS	*piscStatus)
{
	char				   szMsg[1024];
	DWORD				   cbMsg;
	ISC_STATUS			  *piscStatusPtr = piscStatus;
	int	n = 0;

	if (m_fErrorPosted)
		return E_FAIL;
	else
		m_fErrorPosted = TRUE;

	while(isc_interprete(szMsg + n, &piscStatusPtr))
	{
		cbMsg = mbstowcs(
				NULL,
				szMsg,
				sizeof(szMsg));

		if ((!m_pwszDesc) ||
			(2 + cbMsg + (m_pwszDescEnd - m_pwszDesc) > m_cbDesc))
		{
			WCHAR	*pwszTemp = new WCHAR[m_cbDesc + sizeof(szMsg)];

			if (m_pwszDesc)
			{
				memcpy(pwszTemp, m_pwszDesc, m_cbDesc * sizeof(WCHAR));
				delete[] m_pwszDesc;
			}

			m_pwszDesc = pwszTemp;
			m_pwszDescEnd = m_pwszDesc + m_cbDesc;
			m_cbDesc += sizeof(szMsg);
		}

		if (n == 1)
		{
			*m_pwszDescEnd++ = L'\n';
			m_cbDesc++;
		}

		mbstowcs(m_pwszDescEnd, szMsg, cbMsg + 1);
		m_pwszDescEnd += cbMsg;

		n = 1;
		szMsg[0] = '-';
	}

	return PostError(m_pwszDesc, piscStatus[1], E_FAIL);
}