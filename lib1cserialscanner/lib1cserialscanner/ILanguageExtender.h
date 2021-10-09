/*
*  Project name:    lib1cserialscanner
*
*  File name:       ILanguageExtender.h
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#pragma once

#include <Windows.h>
#include <initguid.h>
#include "IInitDone.h"

#ifndef ILANGUAGEEXTENDER_H
#define ILANGUAGEEXTENDER_H

#define NUM_METHODS 3

DEFINE_GUID(IID_ILanguageExtender, 0xab634003, 0xF13D, 0x11d0, 0xa4, 0x59, 0x00, 0x40, 0x95, 0xe1, 0xda, 0xea);

struct __declspec(novtable) ILanguageExtender : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE RegisterExtensionAs(_Inout_ BSTR* pExtensionName) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNProps(_Inout_ long* plProps) = 0;
    virtual HRESULT STDMETHODCALLTYPE FindProp(_In_ BSTR bstrPropName, _Inout_ long* plPropNum) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPropName(_In_ long lPropNum, _In_ long lPropAlias, _Inout_ BSTR* pbstrPropName) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPropVal(_In_ long lPropNum, _Inout_ VARIANT* pvarPropVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetPropVal(_In_ long lPropNum, _In_ VARIANT* varPropVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsPropReadable(_In_ long lPropNum, _Inout_ BOOL* pboolPropRead) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsPropWritable(_In_ long lPropNum, _Inout_ BOOL* pboolPropWrite) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNMethods(_Inout_ long* plMethods) = 0;
    virtual HRESULT STDMETHODCALLTYPE FindMethod(BSTR bstrMethodName, _Inout_ long* plMethodNum) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetMethodName(_In_ long lMethodNum, _In_ long lMethodAlias, _Inout_ BSTR* pbstrMethodName) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNParams(_In_ long lMethodNum, _Inout_ long* plParams) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetParamDefValue(_In_ long lMethodNum, _In_ long lParamNum, _Inout_ VARIANT* pvarParamDefValue) = 0;
    virtual HRESULT STDMETHODCALLTYPE HasRetVal(_In_ long lMethodNum, _Inout_ BOOL* pboolRetValue) = 0;
    virtual HRESULT STDMETHODCALLTYPE CallAsProc(_In_ long lMethodNum, _Inout_ SAFEARRAY** paParams) = 0;
    virtual HRESULT STDMETHODCALLTYPE CallAsFunc(_In_ long lMethodNum, _Inout_ VARIANT* pvarRetValue, _Inout_ SAFEARRAY** paParams) = 0;
};

class SerialScannerLangExt : public ILanguageExtender
{
protected:
    unsigned long   ulRefcount;
    HANDLE          hDevice;
    HANDLE          hReaderThread;
    IAsyncEvent     *pAppEvent;

public:
    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, LPVOID* ppvObject) override;

    //ILanguageExtender
    HRESULT STDMETHODCALLTYPE RegisterExtensionAs(_Inout_ BSTR* pExtensionName) noexcept override;
    HRESULT STDMETHODCALLTYPE GetNProps(_Inout_ long* plProps) noexcept override;
    HRESULT STDMETHODCALLTYPE FindProp(_In_ BSTR bstrPropName, _Inout_ long* plPropNum) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPropName(_In_ long lPropNum, _In_ long lPropAlias, _Inout_ BSTR* pbstrPropName) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPropVal(_In_ long lPropNum, _Inout_ VARIANT* pvarPropVal) noexcept override;
    HRESULT STDMETHODCALLTYPE SetPropVal(_In_ long lPropNum, _In_ VARIANT* varPropVal) noexcept override;
    HRESULT STDMETHODCALLTYPE IsPropReadable(_In_ long lPropNum, _Inout_ BOOL* pboolPropRead) noexcept override;
    HRESULT STDMETHODCALLTYPE IsPropWritable(_In_ long lPropNum, _Inout_ BOOL* pboolPropWrite) noexcept override;
    HRESULT STDMETHODCALLTYPE GetNMethods(_Inout_ long* plMethods) noexcept override;
    HRESULT STDMETHODCALLTYPE FindMethod(BSTR bstrMethodName, _Inout_ long* plMethodNum) noexcept override;
    HRESULT STDMETHODCALLTYPE GetMethodName(_In_ long lMethodNum, _In_ long lMethodAlias, _Inout_ BSTR* pbstrMethodName) noexcept override;
    HRESULT STDMETHODCALLTYPE GetNParams(_In_ long lMethodNum, _Inout_ long* plParams) noexcept override;
    HRESULT STDMETHODCALLTYPE GetParamDefValue(_In_ long lMethodNum, _In_ long lParamNum, _Inout_ VARIANT* pvarParamDefValue) noexcept override;
    HRESULT STDMETHODCALLTYPE HasRetVal(_In_ long lMethodNum, _Inout_ BOOL* pboolRetValue) noexcept override;
    HRESULT STDMETHODCALLTYPE CallAsProc(_In_ long lMethodNum, _Inout_ SAFEARRAY** paParams) noexcept override;
    HRESULT STDMETHODCALLTYPE CallAsFunc(_In_ long lMethodNum, _Inout_ VARIANT* pvarRetValue, _Inout_ SAFEARRAY** paParams) noexcept override;

    // SerialScannerLangExt
    BOOL Open(const WCHAR *PortName, DWORD BaudRate, BYTE DataBits, BYTE Parity, BYTE StopBits) noexcept;
    BOOL Read(WCHAR *Data) noexcept;
    BOOL Close() noexcept;
    SerialScannerLangExt(_In_ IAsyncEvent* pInterface);
    virtual ~SerialScannerLangExt();
};

#endif ILANGUAGEEXTENDER_H
