/*
*  Project name:    lib1cserialscanner
*
*  File name:       IClassFactory.cpp
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#include <Windows.h>
#include "IClassFactory.h"
#include "IInitDone.h"

ULONG64             g_ServerLock = 0;
CRITICAL_SECTION    g_ServerCS;

ScannerClassFactory::ScannerClassFactory() noexcept
{
    this->ulRefcount = 0;
    OutputDebugStringW(L"ScannerClassFactory::ctor ()\r\n");
}

ScannerClassFactory::~ScannerClassFactory() noexcept
{
    WCHAR   text[256];

    wsprintfW(text, L"ScannerClassFactory::dtor Refcount=%u()\r\n", this->ulRefcount);
    OutputDebugStringW(text);
}

ULONG ScannerClassFactory::AddRef()
{
    OutputDebugStringW(L"ScannerClassFactory::AddRef\r\n");
    return InterlockedIncrement(&ulRefcount);
}

ULONG ScannerClassFactory::Release()
{
    const ULONG refcnt = InterlockedDecrement(&this->ulRefcount);

    OutputDebugStringW(L"ScannerClassFactory::Release\r\n");
    if (refcnt == 0)
    {
        delete this;
    }
    return refcnt;
}

HRESULT STDMETHODCALLTYPE ScannerClassFactory::QueryInterface(REFIID riid, LPVOID* ppvObject)
{
    WCHAR   text[256];

    wsprintfW(text, L"ScannerClassFactory::QueryInterface %X\r\n", riid.Data1);
    OutputDebugStringW(text);

    if (ppvObject == NULL)
        return E_INVALIDARG;

    *ppvObject = NULL;
    if ((riid == IID_IUnknown) || (riid == IID_IClassFactory))
    {
        OutputDebugStringW(L"ScannerClassFactory::QueryInterface OK\r\n");
        *ppvObject = (LPVOID)this;
        this->AddRef();
        return NOERROR;
    }
    return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE ScannerClassFactory::LockServer(BOOL fLock) noexcept
{
    OutputDebugStringW(L"ScannerClassFactory::LockServer\r\n");

    EnterCriticalSection(&g_ServerCS);
    if (fLock)
    {
        if (g_ServerLock < MAXULONG64)
            ++g_ServerLock;
    }
    else
    {
        if (g_ServerLock > 0)
            --g_ServerLock;
    }
    LeaveCriticalSection(&g_ServerCS);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ScannerClassFactory::CreateInstance(
    IUnknown* pUnkOuter, REFIID riid, LPVOID* ppvObject)
{
    WCHAR   text[256] = { 0 };
    HRESULT hr = E_OUTOFMEMORY;

    wsprintfW(text, L"ScannerClassFactory::CreateInstance pUnkOuter:%p, riid:%X\r\n", pUnkOuter, riid.Data1);
    OutputDebugStringW(text);

    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION;

    if (ppvObject == NULL)
        return E_INVALIDARG;

    if (riid == IID_IInitDone)
    {
        SerialScannerInit* v = new SerialScannerInit();
        hr = v->QueryInterface(riid, ppvObject);
        if (FAILED(hr))
            delete v;
    }
    
    return hr;
}