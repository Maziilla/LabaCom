#include "interfaces.h"
#include <windows.h>
#include "iid.h"
#include <iostream>
#include "CoCar.h"
#include "CoCarClassFactory.h"

extern ULONG g_lockCount;
extern ULONG g_objCount;

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	HRESULT hr;
	CoCarClassFactory *pCFact = NULL;

	if (rclsid != CLSID_CoCar)
		return CLASS_E_CLASSNOTAVAILABLE;

	pCFact = new CoCarClassFactory;

	hr = pCFact->QueryInterface(riid, ppv);

	if (FAILED(hr))
		delete pCFact;

	return hr;
}

STDAPI DllCanUnloadNow()
{
	if (g_lockCount == 0 && g_objCount == 0)
	{
		return S_OK;
	}
	else
		return S_FALSE;
}
