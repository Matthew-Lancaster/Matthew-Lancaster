//see myscriptsite.h

#include "stdafx.h"
#include "myscriptsite.h"
#include "SiteManager.h"

#include "Application.h"
extern CApplication* pApplication;
CMyScriptSite::~CMyScriptSite() {

	//HRESULT hr = m_scriptp->SetScriptState(SCRIPTSTATE_DISCONNECTED);
	m_scriptp->Close();
	m_parse->Release();
	m_scriptp->Release();
	m_sitemanage->Release();
}