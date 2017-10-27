#pragma warning(disable:4996)

#include "interfaces.h"
#include <windows.h>
#include "iid.h"
#include <iostream>
#include "CoCar.h"

STDMETHODIMP_(DWORD) CoCar::AddRef()
{
	++m_refCount;
	return m_refCount;
}

STDMETHODIMP_(DWORD) CoCar::Release()
{
	if (--m_refCount == 0)
	{
		delete this;
		return 0;
	}
	else
		return m_refCount;
}

STDMETHODIMP CoCar::QueryInterface(REFIID riid, void** pIFace)
{
	if (riid == IID_IUnknown)
	{
		*pIFace = (IUnknown*)(IEngine*)this;
		MessageBox(NULL, L"Handed out IUnknown", L"QI", MB_OK |
			MB_SETFOREGROUND);
	}

	else if (riid == IID_IEngine)
	{
		*pIFace = (IEngine*)this;
		MessageBox(NULL, L"Handed out IEngine", L"QI", MB_OK |
			MB_SETFOREGROUND);
	}

	else if (riid == IID_IStats)
	{
		*pIFace = (IStats*)this;
		MessageBox(NULL, L"Handed out IStats", L"QI", MB_OK |
			MB_SETFOREGROUND);
	}

	else if (riid == IID_ICreateCar)
	{
		*pIFace = (ICreateCar*)this;
		MessageBox(NULL, L"Handed out ICreateCar", L"QI", MB_OK |
			MB_SETFOREGROUND);
	}
	else
	{
		*pIFace = NULL;
		return E_NOINTERFACE;
	}

	((IUnknown*)(*pIFace))->AddRef();
	return S_OK;
}

BSTR	m_petName; // Инициализация через SysAllocString(), 
				   // удаление — через вызов SysFreeString()
int		m_maxSpeed; // Максимальная скорость
int		m_currSpeed; // Текущая скорость СоСаr

extern DWORD g_objCount;

// Конструктор и деструктор СоСаr
CoCar::CoCar() : m_refCount(0)
{
	m_refCount = 0;
	m_currSpeed = m_maxSpeed = 0;
	m_petName = SysAllocString(L"Default Pet Name");
	++g_objCount;
}

CoCar::~CoCar()
{
	--g_objCount;
	if (m_petName)
		SysFreeString(m_petName);
	MessageBox(NULL, L"CoCar is dead", L"Destructor", MB_OK |
		MB_SETFOREGROUND);
}

// Реализация IEngine
STDMETHODIMP CoCar::SpeedUp()
{
	m_currSpeed += 10;
	return S_OK;
}

STDMETHODIMP CoCar::GetMaxSpeed(int* maxSpeed)
{
	*maxSpeed = m_maxSpeed;
	return S_OK;
}

STDMETHODIMP CoCar::GetCurSpeed(int* curSpeed)
{
	*curSpeed = m_currSpeed;
	return S_OK;
}

// Реализация ICreateCar
STDMETHODIMP CoCar::SetPetName(BSTR petName)
{
	SysReAllocString(&m_petName, petName);
	return S_OK;
}

const int MAX_SPEED = 500;

STDMETHODIMP CoCar::SetMaxSpeed(int maxSp)
{
	if (maxSp < MAX_SPEED)
		m_maxSpeed = maxSp;
	return S_OK;
}

// Возвращает клиенту копию внутреннего буфера BSTR 
STDMETHODIMP CoCar::GetPetName(BSTR* petName)
{
	*petName = SysAllocString(m_petName);
	return S_OK;
}

const INT MAX_NAME_LENGTH = 32;

// Информация о СоСаr помещается в блоки сообщений
STDMETHODIMP CoCar::DisplayStats()
{
	// Need to transfer a BSTR to a char array.
	char buff[MAX_NAME_LENGTH];
	WideCharToMultiByte(CP_ACP, NULL, m_petName, -1, buff,
		MAX_NAME_LENGTH, NULL, NULL);

	wchar_t wtext[MAX_NAME_LENGTH];
	mbstowcs(wtext, buff, strlen(buff) + 1);
	LPWSTR ptr = wtext;
	MessageBox(NULL, ptr, L"Pet Name", MB_OK |
		MB_SETFOREGROUND);
	swprintf(wtext, L"%d", m_maxSpeed);
	MessageBox(NULL, ptr, L"Max Speed", MB_OK |
		MB_SETFOREGROUND);
	memset(buff, 0, sizeof(buff));
	memset(wtext, 0, sizeof(wtext));
	return S_OK;
}
//
//STDMETHODIMP
//CoCar::GetIDsOfNames(
//	REFIID riid,
//	OLECHAR ** rgszNames,
//	UINT cNames,
//	LCID lcid,
//	DISPID * rgDispId)
//{
//	return DispGetIDsOfNames(m_ptinfo, rgszNames, cNames, rgDispId);
//};