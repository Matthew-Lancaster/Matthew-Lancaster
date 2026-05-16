/************************************************************************

eventsink.cpp

Experimental code for attaching to events on external objects.

TODO:
Needs tweaking. Doesnt work with IE events, but does with other objects.
(note IE is known to be strange in this regard).

************************************************************************/


#include "stdafx.h"
#include "eventsink.h"
//#include <AfxPriv.h>

CEventSink::CEventSink()
{
	// prepare our internal settings
	m_RefCount = 1;
	m_pPoint = NULL;
	m_pDispatch = NULL;
	m_pTypeInfo = NULL;
}

// this method us used to connect this event sink to a connection point
BOOL CEventSink::Advise( IConnectionPoint *pPoint )
{
	// try to make the connection
	IUnknown *pnt;
	//this->QueryInterface(IID_IUnknown, (LPVOID*)&pnt);
	HRESULT hResult = pPoint->Advise(this, &m_Cookie );
	if( hResult == S_OK )
	{
		// if it worked then store this point so we can later unadvise
		m_pPoint = pPoint;
		return TRUE;
	} return FALSE;
}

// this method is used to break the connection between this sink and the
// connection point it is attached to
BOOL CEventSink::Unadvise()
{
	//ASSERT( m_pPoint );
	//ASSERT( m_pDispatch );
	//ASSERT( m_pTypeInfo );

	// break the connection
	if( m_pPoint->Unadvise( m_Cookie ) != S_OK ) return FALSE;
	// free the object it was attached to
	m_pPoint->Release();
	m_pTypeInfo->Release();
	m_pDispatch->Release();

	return TRUE;
}

// once the connection is made this function is called to store off needed
// data
BOOL CEventSink::SetData( IActiveScript *pScript, IDispatch *pDisp,
						 ITypeInfo *pInfo, char *Name )
{
	//m_ObjectName = Name;
	strncpy(m_ObjectName, Name, sizeof(m_ObjectName));
	m_pScript = pScript;
	m_pTypeInfo = pInfo;
	m_pDispatch = pDisp;

	// we need to make sure the object we are connected to does not exit
	// until we are ready
	m_pDispatch->AddRef();

	return TRUE;
}

// basic ole stuff
STDMETHODIMP_(ULONG) CEventSink::AddRef()
{
	return ++m_RefCount;
}

// basic ole stuff
STDMETHODIMP_(ULONG) CEventSink::Release()
{
	if( --m_RefCount == 0 ) delete this;
	return m_RefCount;
}

// our query interface is special because it will always say that
// we are the interface that the caller is looking for
STDMETHODIMP CEventSink::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	AddRef();

	// doesn't matter what they are looking for we are it
	*ppvObj = ( void * )this;

	return S_OK;
}

// stub these three methods out they won't ever be called in our special case
STDMETHODIMP CEventSink::GetTypeInfoCount( UINT *pctinfo )
{
	return E_NOTIMPL;
}

STDMETHODIMP CEventSink::GetTypeInfo( UINT iTInfo, LCID lcid,
									 ITypeInfo **ppTInfo )
{
	return E_NOTIMPL;
}

STDMETHODIMP CEventSink::GetIDsOfNames( REFIID riid, LPOLESTR *rgszNames,
									   UINT cNames, LCID lcid,
									   DISPID *rgDispId )
{
	return E_NOTIMPL;
}

// this method takes in the id of a method and returns its name for the
// type info provided
BOOL CEventSink::GetNameOfID( DISPID MethodID, char *Name )
{
	USES_CONVERSION;
	//ASSERT( m_pTypeInfo );

	BSTR TempName;
	UINT Count;

	// try to get the name of the method
	HRESULT Rslt = m_pTypeInfo->GetNames( MethodID, &TempName, 1, &Count );
	if( Rslt == S_OK )
	{
		// we found it so return it
		//Name = TempName;
		strncpy(Name, OLE2A(TempName), 1024);
		return TRUE;
	}

	// we could not find it
	return FALSE;
}

STDMETHODIMP CEventSink::Invoke( DISPID dispIdMember, REFIID riid,
                                        LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                        VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr )
{
    USES_CONVERSION;

    // get the name of the calling event
	char pFunctionName[1024];
	char FunctionName[1024];
    if( !GetNameOfID( dispIdMember, pFunctionName ) ) return E_NOTIMPL;

    // add to it our objects name and a "_"
    //FunctionName = m_ObjectName + "_" + FunctionName;
	wsprintf(FunctionName, "%s_%s", m_ObjectName, pFunctionName);
    // get the dispatch for the scripting language
    IDispatch *pScriptDisp;
	m_pScript->QueryInterface(IID_IDispatch, (void**)&pScriptDisp);
    
	///CComPtr< ITypeInfo > spTypeInfo;
	if(m_pScript->GetScriptDispatch(NULL, &pScriptDisp ) != S_OK )
		return E_NOTIMPL;
	/*
	pScriptDisp->GetTypeInfo( 0, LOCALE_SYSTEM_DEFAULT, &spTypeInfo );

	// get the script name space TYPEATTR
	TYPEATTR *pAttr = NULL;
	spTypeInfo->GetTypeAttr( &pAttr );

	// iterate each function
	long f = 0;
	for( f = 0; f < pAttr->cFuncs; ++f )
	{
		// get the FUNCDESC for this function
		FUNCDESC *pFuncDesc = NULL;
		MEMBERID memid = 0;
		spTypeInfo->GetFuncDesc( f, &pFuncDesc );

		// get the name of this function
		CComBSTR bstrName;
		unsigned int cNames = 0;
		spTypeInfo->GetNames( pFuncDesc->memid, &bstrName, 1, &cNames );

		// store the name in the array
		
		// clean up FUNCDESC
		spTypeInfo->ReleaseFuncDesc( pFuncDesc );
	}
	pScriptDisp->GetTypeInfoCount(&val);
*/
	//if(m_pScript->GetScriptDispatch(L"Application.GR_AfterUpdate", &pScriptDisp ) != S_OK )
	//	return E_NOTIMPL;

    // find the method id for that subroutine
    OLECHAR *pName = ( OLECHAR * )T2COLE( FunctionName );
    DISPID dispid;
    HRESULT Rslt;
	UINT val;
	
	Rslt = pScriptDisp->GetIDsOfNames( IID_NULL, &pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid );

    // if we found the sub we looked for
    if( Rslt == S_OK )
    {
        // call it
        Rslt = pScriptDisp->Invoke( dispid, IID_NULL, 9, DISPATCH_METHOD,
                                pDispParams, pVarResult, pExcepInfo, puArgErr );
    }

    // we are done so free the script language's dispatch
    pScriptDisp->Release();

    // return the results
    return Rslt;
}