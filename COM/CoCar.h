#pragma once

#include "interfaces.h"
#include <windows.h>
#include "iid.h"
#include <iostream>

class CoCar : public IEngine, public ICreateCar, public IStats, public IDispatch
{
public:
	CoCar();
	virtual ~CoCar();

	ITypeInfo* m_ptinfo = NULL;
	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void** pIFace);
	STDMETHODIMP_(DWORD)AddRef();
	STDMETHODIMP_(DWORD)Release();

	// IEngine
	STDMETHODIMP SpeedUp();
	STDMETHODIMP GetMaxSpeed(int* maxSpeed);
	STDMETHODIMP GetCurSpeed(int* curSpeed);

	// IStats
	STDMETHODIMP DisplayStats();
	STDMETHODIMP GetPetName(BSTR* petName);

	// ICreateCar
	STDMETHODIMP SetPetName(BSTR petName);
	STDMETHODIMP SetMaxSpeed(int maxSp);

	STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR  ** rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetTypeInfoCount(UINT * pctinfo);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
private:
	unsigned long m_refCount;
};