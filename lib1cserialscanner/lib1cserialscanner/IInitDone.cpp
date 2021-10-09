/*
*  Project name:    lib1cserialscanner
*
*  File name:       IInitDone.cpp
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#include <Windows.h>
#include "IInitDone.h"
#include "ILanguageExtender.h"

SerialScannerInit::SerialScannerInit() noexcept
{
    this->pConnection = NULL;
    this->ulRefcount = 0;
    OutputDebugStringW(L"SerialScannerInit::ctor ()\r\n");
}

SerialScannerInit::~SerialScannerInit()
{
    WCHAR       text[256] = { 0 };
    IDispatch   *oldp = (IDispatch*)InterlockedExchangePointer((PVOID*)&this->pConnection, NULL);

    if (oldp != NULL)
        oldp->Release();

    wsprintfW(text, L"SerialScannerInit::dtor Refcount=%u, pConnection=%p\r\n", this->ulRefcount, this->pConnection);
    OutputDebugStringW(text);
}

ULONG SerialScannerInit::AddRef()
{
    WCHAR       text[256] = { 0 };
    const ULONG refcnt = InterlockedIncrement(&ulRefcount);

    wsprintfW(text, L"SerialScannerInit::AddRef Refcount=%u\r\n", refcnt);
    OutputDebugStringW(text);

    return refcnt;
}

ULONG SerialScannerInit::Release()
{
    WCHAR       text[256] = { 0 };
    const ULONG refcnt = InterlockedDecrement(&this->ulRefcount);

    wsprintfW(text, L"SerialScannerInit::Release Refcount=%u\r\n", refcnt);
    OutputDebugStringW(text);
    if (refcnt == 0)
    {
        delete this;
    }

    return refcnt;
}

HRESULT STDMETHODCALLTYPE SerialScannerInit::QueryInterface(_In_ REFIID riid, _Out_ LPVOID* ppvObject)
{
    WCHAR           text[256] = { 0 };
    IAsyncEvent     *ias =  NULL;
    HRESULT         r = NULL;

    wsprintfW(text, L"SerialScannerInit::QueryInterface %X\r\n", riid.Data1);
    OutputDebugStringW(text);

    if (ppvObject == NULL)
        return E_INVALIDARG;

    *ppvObject = NULL;
    if ((riid == IID_IUnknown) || (riid == IID_IInitDone))
    {
        OutputDebugStringW(L"SerialScannerInit::QueryInterface OK IID_IInitDone\r\n");
        *ppvObject = (LPVOID)this;
        this->AddRef();
        return NOERROR;
    }

    if (riid == IID_ILanguageExtender)
    {
        if (!this->pConnection)
            return E_OUTOFMEMORY;

        r = this->pConnection->QueryInterface(IID_IAsyncEvent, (LPVOID*)&ias);
        if ((r == S_OK) && (ias != NULL))
        {
            OutputDebugStringW(L"SerialScannerInit::QueryInterface OK IID_ILanguageExtender\r\n");
            SerialScannerLangExt* v = new SerialScannerLangExt(ias);
            *ppvObject = (LPVOID)v;
            v->AddRef();
            ias->Release();
        }
        return NOERROR;
    }

    if (riid == IID_ILocale)
    {
        OutputDebugStringW(L"SerialScannerInit::QueryInterface OK IID_ILocale\r\n");
        SerialScannerLocale* v = new SerialScannerLocale;
        *ppvObject = (LPVOID)v;
        v->AddRef();
        return NOERROR;
    }

    return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE SerialScannerInit::Init(_In_ IDispatch* pBackConnection)
{
    IDispatch       *vd = NULL;

    OutputDebugStringW(L"SerialScannerInit::Init\r\n");

    if (pBackConnection == NULL)
        return S_FALSE;

    pBackConnection->AddRef();
    vd = (IDispatch*)InterlockedExchangePointer((PVOID*)&this->pConnection, pBackConnection);
    if (vd != NULL)
        vd->Release();
  
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerInit::Done(void)
{
    WCHAR       text[256] = { 0 };
    IDispatch   *oldp = (IDispatch*)InterlockedExchangePointer((PVOID*)&this->pConnection, NULL);

    if (oldp != NULL)
        oldp->Release();

    wsprintfW(text, L"SerialScannerInit::Done Refcount=%u, pConnection=%p\r\n", this->ulRefcount, oldp);
    OutputDebugStringW(text);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerInit::GetInfo(_Inout_ SAFEARRAY** pInfo) noexcept
{
    long        lInd = 0;
    VARIANT     varVersion = { 0 };

    OutputDebugStringW(L"SerialScannerInit::GetInfo\r\n");

    if (pInfo == NULL)
        return E_FAIL;

    V_VT(&varVersion) = VT_I4;
    V_I4(&varVersion) = 2000;
    return SafeArrayPutElement(*pInfo, &lInd, &varVersion);
}

SerialScannerLocale::SerialScannerLocale() noexcept
{
    this->ulRefcount = 0;
    OutputDebugStringW(L"SerialScannerLocale::ctor ()\r\n");
}

SerialScannerLocale::~SerialScannerLocale() noexcept
{
    WCHAR   text[256];

    wsprintfW(text, L"SerialScannerLocale::dtor Refcount=%u()\r\n", this->ulRefcount);
    OutputDebugStringW(text);
}

ULONG SerialScannerLocale::AddRef()
{
    OutputDebugStringW(L"SerialScannerLocale::AddRef\r\n");
    return InterlockedIncrement(&ulRefcount);
}

ULONG SerialScannerLocale::Release()
{
    const ULONG refcnt = InterlockedDecrement(&this->ulRefcount);

    OutputDebugStringW(L"SerialScannerLocale::Release\r\n");
    if (refcnt == 0)
    {
        delete this;
    }
    return refcnt;
}

HRESULT STDMETHODCALLTYPE SerialScannerLocale::QueryInterface(_In_ REFIID riid, _Out_ LPVOID* ppvObject)
{
    WCHAR   text[256];

    wsprintfW(text, L"SerialScannerLocale::QueryInterface %X\r\n", riid.Data1);
    OutputDebugStringW(text);

    if (ppvObject == NULL)
        return E_INVALIDARG;

    *ppvObject = NULL;
    if ((riid == IID_IUnknown) || (riid == IID_ILocale))
    {
        OutputDebugStringW(L"SerialScannerLocale::QueryInterface OK\r\n");
        *ppvObject = (LPVOID)this;
        this->AddRef();
        return NOERROR;
    }

    return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLocale::SetLocale(_In_ BSTR bstrLocale) noexcept
{
    WCHAR   text[256];

    wsprintfW(text, L"SerialScannerInit::SetLocale %s\r\n", bstrLocale);
    OutputDebugStringW(text);
    return S_OK;
}