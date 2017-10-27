#pragma once

#include "interfaces.h"
#include <windows.h>
#include "iid.h"
#include <iostream>

class CoCar : public IEngine, public ICreateCar, public IStats
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

	//STDMETHODIMP GetIDsOfNames(
	//	[in]                   REFIID   riid,
	//	[in, size_is(cNames)]  LPOLESTR *rgszNames,
	//	[in]                   UINT     cNames,
	//	[in]                   LCID     lcid,
	//	[out, size_is(cNames)] DISPID   *rgDispId
	//);
	//STDMETHODIMP GetTypeInfo(
	//	[in]  UINT      iTInfo,
	//	[in]  LCID      lcid,
	//	[out] ITypeInfo **ppTInfo
	//);
	//STDMETHODIMP GetTypeInfoCount(
	//	[out] UINT *pctinfo
	//);
	//STDMETHODIMP Invoke(
	//	[in]      DISPID     dispIdMember,
	//	[in]      REFIID     riid,
	//	[in]      LCID       lcid,
	//	[in]      WORD       wFlags,
	//	[in, out] DISPPARAMS *pDispParams,
	//	[out]     VARIANT    *pVarResult,
	//	[out]     EXCEPINFO  *pExcepInfo,
	//	[out]     UINT       *puArgErr
	//);
private:
	unsigned long m_refCount;
};