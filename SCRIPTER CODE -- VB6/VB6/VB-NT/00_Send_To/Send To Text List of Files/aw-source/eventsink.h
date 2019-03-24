#ifndef EVENTSINKCLASS
#define EVENTSINKCLASS

#pragma once
#include "resource.h"       // main symbols

#include "gen_activewa.h"
#include "MyScriptSite.h"


class CEventSink : public IDispatch
{
public:

	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();
	STDMETHOD(QueryInterface)(REFIID, LPVOID*);

	STDMETHOD(GetTypeInfoCount)(UINT*);
	STDMETHOD(GetTypeInfo)(UINT, LCID, LPTYPEINFO*);
	STDMETHOD(GetIDsOfNames)(REFIID, LPOLESTR*, UINT, LCID, DISPID*);
	STDMETHOD(Invoke)(DISPID, REFIID, LCID, WORD, DISPPARAMS*, LPVARIANT,LPEXCEPINFO, UINT*);

	long m_RefCount;

	IActiveScript *m_pScript;
	ITypeInfo *m_pTypeInfo;
	IConnectionPoint *m_pPoint;
	DWORD m_Cookie;
	IDispatch *m_pDispatch;
	char m_ObjectName[1024];

	BOOL GetNameOfID( DISPID MethodID, char* Name );
	BOOL SetData( IActiveScript*pScript, IDispatch *pDisp,	ITypeInfo *pTypeInfo, char *Name );
	BOOL Advise( IConnectionPoint *pPoint );
	BOOL Unadvise();
	
	CEventSink();
};

#endif
