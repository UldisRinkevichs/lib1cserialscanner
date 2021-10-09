/*
*  Project name:    lib1cserialscanner
*
*  File name:       IInitDone.h
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#pragma once

#include <Windows.h>
#include <initguid.h>

#ifndef IINITDONE_H
#define IINITDONE_H

DEFINE_GUID(IID_IInitDone, 0xab634001, 0xF13D, 0x11d0, 0xa4, 0x59, 0x00, 0x40, 0x95, 0xe1, 0xda, 0xea);
DEFINE_GUID(IID_IAsyncEvent, 0xab634004, 0xF13D, 0x11d0, 0xa4, 0x59, 0x00, 0x40, 0x95, 0xe1, 0xda, 0xea);
DEFINE_GUID(IID_ILocale, 0xe88a191e, 0x8f52, 0x4261, 0x9f, 0xae, 0xff, 0x7a, 0xa8, 0x4f, 0x5d, 0x7e);

struct __declspec(novtable) IInitDone : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Init(_In_ IDispatch* pBackConnection) = 0;
    virtual HRESULT STDMETHODCALLTYPE Done(void) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetInfo(_Inout_ SAFEARRAY** pInfo) = 0;
};

struct __declspec(novtable) ILocale : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetLocale(_In_ BSTR bstrLocale) = 0;
};

MIDL_INTERFACE("ab634004-f13d-11d0-a459-004095e1daea")
IAsyncEvent: public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetEventBufferDepth(long lDepth) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetEventBufferDepth(long* plDepth) = 0;
    virtual HRESULT STDMETHODCALLTYPE ExternalEvent(BSTR bstrSource, BSTR bstrMessage, BSTR bstrData) = 0;
    virtual HRESULT STDMETHODCALLTYPE CleanBuffer(void) = 0;
};

class SerialScannerInit : public IInitDone
{
protected:
    unsigned long   ulRefcount;
    IDispatch       *pConnection;

public:
    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE QueryInterface(_In_ REFIID riid, _Out_ LPVOID* ppvObject) override;

    // IInitDone
    HRESULT STDMETHODCALLTYPE Init(_In_ IDispatch* pBackConnection) override;
    HRESULT STDMETHODCALLTYPE Done(void) override;
    HRESULT STDMETHODCALLTYPE GetInfo(_Inout_ SAFEARRAY** pInfo) noexcept override;

    // SerialScannerInit
    SerialScannerInit() noexcept;
    virtual ~SerialScannerInit();
};

class SerialScannerLocale : public ILocale
{
protected:
    unsigned long   ulRefcount;

public:
    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE QueryInterface(_In_ REFIID riid, _Out_ LPVOID* ppvObject) override;

    // ILocale
    HRESULT STDMETHODCALLTYPE SetLocale(_In_ BSTR bstrLocale) noexcept override;

    // SerialScannerLocale
    SerialScannerLocale() noexcept;
    virtual ~SerialScannerLocale() noexcept;
};

#endif /* IINITDONE_H */
