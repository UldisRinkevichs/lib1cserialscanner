/*
*  Project name:    lib1cserialscanner
*
*  File name:       IClassFactory.h
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#pragma once

#include <Windows.h>
#include <initguid.h>

#ifndef ICLASSFACTORY_H
#define ICLASSFACTORY_H

DEFINE_GUID(CLSID_lib1CSerialScanner, 0x5c344952, 0x17b6, 0x4fc5, 0x9b, 0xfc, 0xd8, 0xb4, 0x17, 0x6e, 0x10, 0x18);

extern CRITICAL_SECTION     g_ServerCS;
extern ULONG64              g_ServerLock;

class ScannerClassFactory : public IClassFactory
{
protected:
    unsigned long   ulRefcount;

public:
    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, LPVOID* ppvObject) override;

    // IClassFactory
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, LPVOID* ppvObject) override;

    // ScannerClassFactory
    ScannerClassFactory() noexcept;
    virtual ~ScannerClassFactory() noexcept;
};

#endif /* ICLASSFACTORY_H */