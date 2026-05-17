/************************************************************************

myscriptsite.h/.cpp

The IActiveScriptSite implementation for internal script hosting.

Objects of this class essentially represent running hosted scripts.

They are double linked with the CSiteManager due to the need to share
info between objects. This implementation could probably be improved.

TODO:
Fix quitting process.
Rename to something sensible. (CVS issue)

************************************************************************/

#pragma once

#ifndef SCRIPTSITE_H
#define SCRIPTSITE_H

#include <windows.h>
#include <activscp.h>
#include "gen.h"
#include <string>

extern winampGeneralPurposePlugin plugin;
extern long quit_ipc;

class CSiteManager;

class CMyScriptSite : public IActiveScriptSite, public IActiveScriptSiteWindow {
private:
	ULONG m_dwRef;
public:
	IUnknown *m_pUnkScriptObject;
	IUnknown *m_pUnkManagerObject;
	ITypeInfo *m_pTypeInfo;
	
	IActiveScript* m_scriptp;
	IActiveScriptParse* m_parse;
	
	CSiteManager *m_sitemanage;

	std::string m_desc;
	std::string *m_arguments;
	DWORD threadid;

	CMyScriptSite::CMyScriptSite() {
		m_pUnkScriptObject = 0;
		m_pUnkManagerObject = 0;
		m_pTypeInfo = 0;
		m_dwRef = 1;
		threadid = GetCurrentThreadId();
	}
	CMyScriptSite::~CMyScriptSite();

	// IUnknown methods...
	virtual HRESULT __stdcall QueryInterface(REFIID riid,
		void **ppvObject) {
			//*ppvObject = NULL;
			if (riid == IID_IActiveScriptSiteWindow)
			{
				*ppvObject = static_cast<IActiveScriptSiteWindow*>(this);
				//*ppvObject = this;
				return S_OK;
			}
			else
				*ppvObject = NULL;

			return E_NOTIMPL;
		}
		virtual ULONG _stdcall AddRef(void) {
			return ++m_dwRef;
		}
		virtual ULONG _stdcall Release(void) {
			if(--m_dwRef == 0) return 0;
			return m_dwRef;
		}

		// IActiveScriptSite methods...
		virtual HRESULT __stdcall GetLCID(LCID *plcid) {
			return S_OK;
		}

		virtual HRESULT __stdcall GetItemInfo(LPCOLESTR pstrName,DWORD dwReturnMask, IUnknown **ppunkItem, ITypeInfo **ppti) 
		{
			// Is it expecting an ITypeInfo?
			if(ppti) {
				// Default to NULL.
				*ppti = NULL;

				// See if asking about ITypeInfo... 
				if(dwReturnMask & SCRIPTINFO_ITYPEINFO) {
					*ppti = m_pTypeInfo;
				}
			}

			// Is the engine passing an IUnknown buffer?
			if(ppunkItem) {
				// Default to NULL.
				*ppunkItem = NULL;

				// Is Script Engine looking for an IUnknown for our object?
				if(dwReturnMask & SCRIPTINFO_IUNKNOWN) {
					// Check for our object name...
					if (!_wcsicmp(L"Application", pstrName)) {
						// Provide our object.
						*ppunkItem = m_pUnkScriptObject;
						// Addref our object...
						m_pUnkScriptObject->AddRef();
					}

					if (!_wcsicmp(L"SiteManager", pstrName)) {
						// Provide our object.
						*ppunkItem = m_pUnkManagerObject;
						// Addref our object...
						m_pUnkManagerObject->AddRef();
					}
				}
			}

			return S_OK;
		}

		virtual HRESULT __stdcall GetDocVersionString(BSTR *pbstrVersion) {
			return S_OK;
		}

		virtual HRESULT __stdcall OnScriptTerminate(const VARIANT *pvarResult, const EXCEPINFO *pexcepInfo)
		{
			return S_OK;
		}

		virtual HRESULT __stdcall OnStateChange(SCRIPTSTATE ssScriptState) 
		{
			return S_OK;
		}

		virtual HRESULT __stdcall OnScriptError(IActiveScriptError *pse)
		{
			USES_CONVERSION;

			EXCEPINFO ei;
			DWORD     dwSrcContext;
			ULONG     ulLine;
			LONG      ichError;
			BSTR      bstrLine = NULL;
			char	  errmsg[1024];

			pse->GetExceptionInfo(&ei);
			pse->GetSourcePosition(&dwSrcContext, &ulLine, &ichError);
			pse->GetSourceLineText(&bstrLine);

			if (ei.scode == 0x80004005)
			{
				//Script was interrupted, terminate without error msg
				//Sends a WM_QUIT to the workerthread so it knows to quit.
				//This call is kind of redundant, as the main wnd proc already sends this message
				PostThreadMessage(threadid, WM_QUIT, 0, 0);	
			} else{
				wsprintf(errmsg, "Line %d\n%s\n%s", ulLine, OLE2T(bstrLine), OLE2T(ei.bstrDescription));
				MessageBox(NULL, errmsg, "Error", MB_SETFOREGROUND | MB_TASKMODAL);
				PostMessage(plugin.hwndParent, quit_ipc, (WPARAM)this, 0);
				//PostThreadMessage(threadid, WM_QUIT, 0, 0);	
			}

			SysFreeString(ei.bstrSource);
			SysFreeString(ei.bstrDescription);
			SysFreeString(ei.bstrHelpFile);

			return S_OK;
		}

		virtual HRESULT __stdcall OnEnterScript(void) {
			return S_OK;
		}

		virtual HRESULT __stdcall OnLeaveScript(void) {
			return S_OK;
		}

		//SiteWindow
		virtual HRESULT __stdcall GetWindow(
			HWND *phwnd  // address of variable for window handle
			)
		{
			*phwnd  = plugin.hwndParent;
			return S_OK;
		}

		virtual HRESULT __stdcall EnableModeless(
			BOOL fEnable  // enable flag
			)
		{
			return S_OK;
		}

};

#endif