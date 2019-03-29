#pragma once

#ifndef SITEMANAGE_H
#define SITEMANAGE_H

#include "resource.h"       // main symbols
#include "eventsink.h"
#include "gen_activewa.h"


class CMyScriptSite;// : public IActiveScriptSite, public IActiveScriptSiteWindow;

// CSiteManager

class ATL_NO_VTABLE CSiteManager : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CSiteManager, &CLSID_SiteManager>,
	public IDispatchImpl<ISiteManager, &IID_ISiteManager, &LIBID_ActiveWinamp, /*wMajor =*/ 0xffff, /*wMinor =*/ 0xffff>
{
public:
	CSiteManager()
	{
	}

DECLARE_NOT_AGGREGATABLE(CSiteManager)

BEGIN_COM_MAP(CSiteManager)
	COM_INTERFACE_ENTRY(ISiteManager)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	CMyScriptSite *m_hostSite;
	STDMETHOD(AttachEvents)(IDispatch* EventObject, BSTR ObjectPrefix);
	BOOL GetEventTypeInfo( IDispatch *pDisp, ITypeInfo *&pResultTypeInfo );
	STDMETHOD(Quit)(void);
	STDMETHOD(get_Description)(BSTR* pVal);
	STDMETHOD(put_Description)(BSTR newVal);
	STDMETHOD(get_Arguments)(BSTR* pVal);
};

#endif