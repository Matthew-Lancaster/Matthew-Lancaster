/************************************************************************

SiteManager.cpp

The actual COM object implementation for the SiteManager object.
This object is only exposed to internal scripts, as that is the only
place it makes sense to use.

It lets a script manage things like its description, when to quit, script
arguments and attaching to events.

TODO:
Keep a list of EventSinks that are attached, so that they can be cleaned
up on quit.
Fix "quit" implementation to something better.

************************************************************************/

#include "stdafx.h"
#include "SiteManager.h"
#include ".\sitemanager.h"

extern HWND HwndPl;
extern long quit_ipc;

// CSiteManager
//This method leaks: AW should keep a std list of event sinks and clean up.
//Its not a problem as long as this method isn't called.
//Its really only needed for internal scripting to attach to external object
//event handlers.
STDMETHODIMP CSiteManager::AttachEvents(IDispatch* EventObject, BSTR ObjectPrefix)
{
	USES_CONVERSION;
	// First we need to get the type info ( ITypeInfo ) from the object
	// that was passed to us ( this function is defined below in this
	// example
	ITypeInfo *pTypeInfo;
	if( !GetEventTypeInfo(EventObject, pTypeInfo ) ) return FALSE;

	// Next we need to get the IID of the event interface of the object
	// that was passed to us
	IID EventIID;
	TYPEATTR *pEventAttr;
	if( pTypeInfo->GetTypeAttr( &pEventAttr ) == S_OK )
	{
		//ASSERT( pEventAttr != NULL );
		// get the IID
		EventIID = pEventAttr->guid;
		// now make sure we free the TYPEATTR that the ITypeInfo
		// allocated for us
		pTypeInfo->ReleaseTypeAttr( pEventAttr );
	}
	else
	{
		// we failed so release the type info and kick out
		pTypeInfo->Release();
		return FALSE;
	}

	// now that we know the events IID we can get the event connection
	// point container from it
	IConnectionPointContainer *pContainer;
	HRESULT hResult = EventObject->QueryInterface( IID_IConnectionPointContainer,
		( void ** )&pContainer );
	// if we could not get a connection point container then there is
	// something wrong and we need to clean up and tell the user
	if( !hResult == S_OK )
	{
		pTypeInfo->Release();
		return FALSE;
	}

	// once you have the container find the right connection point
	IConnectionPoint *pPoint;
	hResult = pContainer->FindConnectionPoint( EventIID, &pPoint );
	// and now we are done with the container so we can free it up
	pContainer->Release();

	// we could not find the connection point so clean up and exit
	if( !hResult == S_OK )
	{
		pTypeInfo->Release();
		return FALSE;
	}

	// create a sink
	CEventSink *pSink = new CEventSink;
	//ASSERT( pSink );

	// attach it to the connection point
	if( pSink->Advise( pPoint ) )
	{
		// if it worked store the data it needs
		//CComQIPtr<IDispatch> pDisp(MeObject);
		pSink->SetData(m_hostSite->m_scriptp , EventObject, pTypeInfo, OLE2A(ObjectPrefix));
		// add it to a list so at script exit time it can have
		// the Unadvise method called on it and be release
		//m_EventSinks.Add( pSink );
	}
	else
	{
		// the connection failed so clean up
		pTypeInfo->Release();
		pPoint->Release();
		pSink->Release();
	}

	// let the caller know if everything worked
	return ( hResult == S_OK );
}


#define IMPLTYPE_MASK \
	(IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE | IMPLTYPEFLAG_FRESTRICTED)


#define IMPLTYPE_DEFAULTSOURCE \
	(IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)

BOOL CSiteManager::GetEventTypeInfo( IDispatch *pDisp,
								ITypeInfo *&pResultTypeInfo )
{
	//ASSERT( pDisp );
	pResultTypeInfo = NULL;

	// first get the IProvideClassInfo interface
	IProvideClassInfo *pPCI = NULL;
	ITypeInfo *pClassInfo = NULL;
	TYPEATTR *pClassAttr;

	if( pDisp->QueryInterface( IID_IProvideClassInfo, ( LPVOID * )&pPCI )
		== S_OK )
	{
		pPCI->GetClassInfo( &pClassInfo );
	}
	else
	{
		ITypeInfo *pTypeInfo = NULL;
		ITypeLib *pTlib = NULL;

		pDisp->GetTypeInfo( 0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo );
		TYPEATTR *pAttr = NULL;
		pTypeInfo->GetTypeAttr( &pAttr );
		GUID mguid = pAttr->guid;
		UINT ind = 0;
		pTypeInfo->GetContainingTypeLib(&pTlib, &ind);
		UINT tc = pTlib->GetTypeInfoCount();

		pTypeInfo->ReleaseTypeAttr(pAttr);
		pTypeInfo->Release();


		ITypeInfo *pTinfo = NULL;
		TYPEATTR *pAttr2 = NULL;
		TYPEATTR *pAttr3 = NULL;
		TYPEKIND Tkind;

		for (int k = 0; k < tc; k++)
		{
			pTlib->GetTypeInfoType(k, &Tkind);
			if (Tkind == TKIND_COCLASS)
			{
				pTlib->GetTypeInfo(k, &pClassInfo);
				pClassInfo->GetTypeAttr( &pAttr2 );

				int nFlags;
				HREFTYPE hRefType;
				for(unsigned int i = 0; i < pAttr2->cImplTypes; i++ )
				{
					pClassInfo->GetImplTypeFlags(i, &nFlags);

					if ( (nFlags & IMPLTYPE_MASK ) == IMPLTYPEFLAG_FDEFAULT  )
					{
						pTinfo = NULL;
						pClassInfo->GetRefTypeOfImplType( i, &hRefType);
						pClassInfo->GetRefTypeInfo(hRefType, &pTinfo );

						pTinfo->GetTypeAttr(&pAttr3);
						if (IsEqualGUID(mguid, pAttr3->guid))
						{
							//printf("Number of functions: %d\n", pAttr3->cFuncs);
							pTinfo->ReleaseTypeAttr(pAttr3);
							pTinfo->Release();
							break;
						}
						pTinfo->ReleaseTypeAttr(pAttr3);
						pTinfo->Release();
					}
				}
				pClassInfo->ReleaseTypeAttr(pAttr2);
				if (ind == 100)
					break;
				pClassInfo->Release();
			}
			if (ind == 100)
				break;
		}

		pTlib->Release();

	}

	if (pClassInfo != NULL)
	{
		pClassInfo->GetTypeAttr( &pClassAttr );
		int nFlags;
		HREFTYPE hRefType;
		for( unsigned int i = 0; i < pClassAttr->cImplTypes;
			i++ )
		{
			if( pClassInfo->GetImplTypeFlags( i, &nFlags )
				== S_OK && ( ( nFlags & IMPLTYPE_MASK )
				== IMPLTYPE_DEFAULTSOURCE ) )
			{
				pResultTypeInfo = NULL;

				// we fount the it now we need to get
				// the type info interface that we will
				// return to the caller
				if( pClassInfo->GetRefTypeOfImplType( i,
					&hRefType) == S_OK )
				{
					pClassInfo->GetRefTypeInfo(hRefType, &pResultTypeInfo );
					//pResultTypeInfo->GetTypeAttr(&pAttr3);
					//pResultTypeInfo->ReleaseTypeAttr(pAttr3);
					//ASSERT( pResultTypeInfo != NULL );
				}
				break;
			}
		}
		
		if (pPCI != NULL)
			pPCI->Release();
		
		pClassInfo->ReleaseTypeAttr(pClassAttr);
		pClassInfo->Release();
	}

	// return the results
	return pResultTypeInfo ? TRUE : FALSE;
}
STDMETHODIMP CSiteManager::Quit(void)
{
	//Send the quit script IPC to the main window with the script object to terminate
	SendMessage(plugin.hwndParent, quit_ipc, (WPARAM)m_hostSite, (LPARAM)0);
	return S_OK;
}

STDMETHODIMP CSiteManager::get_Description(BSTR* pVal)
{
	USES_CONVERSION;
	*pVal = ::SysAllocString(T2OLE(m_hostSite->m_desc.c_str()));
	return S_OK;
}

STDMETHODIMP CSiteManager::put_Description(BSTR newVal)
{
	USES_CONVERSION;
	m_hostSite->m_desc = OLE2T(newVal);
	return S_OK;
}

STDMETHODIMP CSiteManager::get_Arguments(BSTR* pVal)
{
	USES_CONVERSION;
	*pVal = ::SysAllocString(T2OLE(m_hostSite->m_arguments->c_str()));
	return S_OK;
}
