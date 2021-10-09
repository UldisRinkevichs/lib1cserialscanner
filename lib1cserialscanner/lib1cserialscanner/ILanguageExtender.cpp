/*
*  Project name:    lib1cserialscanner
*
*  File name:       ILanguageExtender.cpp
*  Created on:      Sep 30, 2021
*  Modified on:     Oct 3, 2021
*
*/

#include <Windows.h>
#include <comutil.h>
#include "ILanguageExtender.h"

#define COM_BUFFER_SIZE     2048

typedef struct _RTHREADPARAMS {
    HANDLE          hDevice;
    HANDLE          hMLock;
    IAsyncEvent     *pAppEvent;
} RTHREADPARAMS, *PRTHREADPARAMS;

const WCHAR* SerialScannerMethodNames[NUM_METHODS] =
{
    L"Open",
    L"Read",
    L"Close"
};

const int SerialScannerMethodParamsCount[NUM_METHODS] =
{
    5, 1, 0
};

HRESULT GetNParam(SAFEARRAY *parray, long index, VARIANT* pv) noexcept
{
    if ((!parray) || (!pv))
        return E_INVALIDARG;

    if (!((parray)->fFeatures & FADF_VARIANT))
        return E_INVALIDARG;

    return SafeArrayGetElement(parray, &index, pv);
}

DWORD WINAPI SerialScannerReaderThread(_In_ LPVOID lpParameter)
{
    char                    cv = '\0', text[COM_BUFFER_SIZE] = {0};
    WCHAR                   wtext[COM_BUFFER_SIZE] = { 0 };
    PRTHREADPARAMS          v = (PRTHREADPARAMS)lpParameter;
    HANDLE                  f = NULL;
    DWORD                   chars_read = 0, p = 0;
    IAsyncEvent             *pAppEvent = NULL;
    BSTR                    s_src= NULL, s_msg = NULL, s_data = NULL;

    if (v == NULL)
        return 1;

    f = v->hDevice;
    pAppEvent = v->pAppEvent;
    SetEvent(v->hMLock);

    if (pAppEvent == NULL)
        return 0;

    OutputDebugStringW(L"SerialScannerReaderThread - start\r\n");

    while (ReadFile(f, &cv, 1, &chars_read, NULL))
    {
        if (cv == '\r')
        {
            text[p] = '\0';

            memset(wtext, 0, sizeof(wtext));
            if (MultiByteToWideChar(CP_UTF8, 0, text, p, wtext, COM_BUFFER_SIZE-1))
            {
                s_src = SysAllocString(L"SerialScanner");
                s_msg = SysAllocString(L"CODE");
                s_data = SysAllocString(wtext);
                pAppEvent->ExternalEvent(s_src, s_msg, s_data);
                SysFreeString(s_src);
                SysFreeString(s_msg);
                SysFreeString(s_data);
            }

            p = 0;
            continue;
        }

        text[p] = cv;
        ++p;

        if (p >= (COM_BUFFER_SIZE-1))
            p = 0;
    }

    OutputDebugStringW(L"SerialScannerReaderThread - exit\r\n");
    return 0;
}

BOOL SerialScannerLangExt::Open(const WCHAR* PortName, DWORD BaudRate, BYTE DataBits, BYTE Parity, BYTE StopBits) noexcept
{
    HANDLE          f = NULL;
    DCB             com_stat = { 0 };
    WCHAR           text[256] = { 0 };
    DWORD           tid = 0;
    RTHREADPARAMS   tp = { 0 };

    wsprintfW(text, L"SerialScannerLangExt::Open(\"%s\", %u, %u, %u, %u)\r\n", PortName, BaudRate, DataBits, Parity, StopBits);
    OutputDebugStringW(text);

    if ((this->hDevice != NULL) || (PortName == NULL))
        return FALSE;

    f = CreateFileW(PortName, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (f == INVALID_HANDLE_VALUE)
        return FALSE;

    com_stat.DCBlength = sizeof(com_stat);
    if (GetCommState(f, &com_stat))
    {
        com_stat.BaudRate = BaudRate;
        com_stat.ByteSize = DataBits;
        com_stat.Parity = Parity;
        com_stat.StopBits = StopBits;
        com_stat.EvtChar = '\n';
        com_stat.fParity = TRUE;

        SetCommState(f, &com_stat);
    }

    tp.hDevice = f;
    tp.pAppEvent = this->pAppEvent;
    tp.hMLock = CreateEventExW(NULL, NULL, 0, EVENT_ALL_ACCESS);
    if (tp.hMLock)
    {
        wsprintfW(text, L"SerialScannerLangExt::Open - CreateThread: %p,%p,%p\r\n", tp.hDevice, tp.pAppEvent, tp.hMLock);
        OutputDebugStringW(text);
        this->hReaderThread = CreateThread(NULL, 0, &SerialScannerReaderThread, (LPVOID)&tp, 0, &tid);
        
        wsprintfW(text, L"SerialScannerLangExt::Open - CreateThread: %p\r\n", this->hReaderThread);
        OutputDebugStringW(text);
        
        WaitForSingleObjectEx(tp.hMLock, INFINITE, FALSE);
        CloseHandle(tp.hMLock);
    }

    this->hDevice = f;
    OutputDebugStringW(L"SerialScannerLangExt::Open\r\n");
    return TRUE;
}

BOOL SerialScannerLangExt::Read(WCHAR* Data) noexcept
{
    UNREFERENCED_PARAMETER(Data);
    return TRUE;
}

BOOL SerialScannerLangExt::Close() noexcept
{
    DWORD   r = 0;
    WCHAR   text[256] = { 0 };
    HANDLE
        hthread = InterlockedExchangePointer(&this->hReaderThread, NULL),
        hdev = InterlockedExchangePointer(&this->hDevice, NULL);

    OutputDebugStringW(L"SerialScannerLangExt::Close() 1\r\n");
    
    if (hdev != NULL)
    {
        CancelIo(hdev);
        CancelSynchronousIo(hthread);
        CloseHandle(hdev);
    }

    OutputDebugStringW(L"SerialScannerLangExt::Close() 2\r\n");

    if (hthread != NULL)
    {
        OutputDebugStringW(L"SerialScannerLangExt::Close - Thread wait ...\r\n");
        r = WaitForSingleObjectEx(hthread, 1000, FALSE);
        if (WAIT_OBJECT_0 != r)
        {
            OutputDebugStringW(L"SerialScannerLangExt::Close - Thread wait failed or timed out.\r\n");
            TerminateThread(hthread, 1);
        }
        wsprintfW(text, L"SerialScannerLangExt::Close - wait status:%X\r\n", r);
        OutputDebugStringW(text);
        CloseHandle(hthread);
    }

    return FALSE;
}

SerialScannerLangExt::SerialScannerLangExt(_In_ IAsyncEvent* pInterface)
{
    WCHAR           text[256] = { 0 };

    this->pAppEvent = pInterface;
    this->ulRefcount = 0;
    this->hDevice = NULL;
    this->hReaderThread = NULL; 
    wsprintfW(text, L"SerialScannerLangExt::ctor %p\r\n", pInterface);
    OutputDebugStringW(text);
}

SerialScannerLangExt::~SerialScannerLangExt()
{
    WCHAR           text[256] = { 0 };

    this->Close();
    wsprintfW(text, L"SerialScannerLangExt:dtor Refcount=%u\r\n", this->ulRefcount);
    OutputDebugStringW(text);
}

ULONG SerialScannerLangExt::AddRef()
{
    OutputDebugStringW(L"SerialScannerLangExt::AddRef\r\n");
    return InterlockedIncrement(&ulRefcount);
}

ULONG SerialScannerLangExt::Release()
{
    const ULONG refcnt = InterlockedDecrement(&this->ulRefcount);

    OutputDebugStringW(L"SerialScannerLangExt::Release\r\n");
    if (refcnt == 0)
    {
        delete this;
    }
    return refcnt;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::QueryInterface(REFIID riid, LPVOID* ppvObject)
{
    WCHAR   text[256];

    wsprintfW(text, L"SerialScannerLangExt::QueryInterface %X\r\n", riid.Data1);
    OutputDebugStringW(text);

    if (ppvObject == NULL)
        return E_INVALIDARG;

    *ppvObject = NULL;
    if ((riid == IID_IUnknown) || (riid == IID_ILanguageExtender))
    {
        OutputDebugStringW(L"SerialScannerLangExt::QueryInterface OK\r\n");
        *ppvObject = (LPVOID)this;
        this->AddRef();
        return NOERROR;
    }
    return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::RegisterExtensionAs(_Inout_ BSTR* pExtensionName) noexcept
{
    OutputDebugStringW(L"SerialScannerLangExt::RegisterExtensionAs\r\n");

    if (pExtensionName != NULL)
    {
        *pExtensionName = SysAllocString(L"SerialScanner");
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetNProps(_Inout_ long* plProps) noexcept
{
    OutputDebugStringW(L"SerialScannerLangExt::GetNProps\r\n");

    if (plProps != NULL)
    {
        *plProps = 0;
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::FindProp(_In_ BSTR bstrPropName, _Inout_ long* plPropNum) noexcept
{
    UNREFERENCED_PARAMETER(bstrPropName);

    OutputDebugStringW(L"SerialScannerLangExt::FindProp\r\n");
    if (plPropNum != NULL)
    {
        *plPropNum = -1;
    }

    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetPropName(_In_ long lPropNum, _In_ long lPropAlias, _Inout_ BSTR* pbstrPropName) noexcept
{
    UNREFERENCED_PARAMETER(lPropNum);
    UNREFERENCED_PARAMETER(lPropAlias);
    UNREFERENCED_PARAMETER(pbstrPropName);
    OutputDebugStringW(L"SerialScannerLangExt::GetPropName\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetPropVal(_In_ long lPropNum, _Inout_ VARIANT* pvarPropVal) noexcept
{
    UNREFERENCED_PARAMETER(lPropNum);
    UNREFERENCED_PARAMETER(pvarPropVal);
    OutputDebugStringW(L"SerialScannerLangExt::GetPropVal\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::SetPropVal(_In_ long lPropNum, _In_ VARIANT* varPropVal) noexcept
{
    UNREFERENCED_PARAMETER(lPropNum);
    UNREFERENCED_PARAMETER(varPropVal);
    OutputDebugStringW(L"SerialScannerLangExt::SetPropVal\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::IsPropReadable(_In_ long lPropNum, _Inout_ BOOL* pboolPropRead) noexcept
{
    UNREFERENCED_PARAMETER(lPropNum);
    UNREFERENCED_PARAMETER(pboolPropRead);
    OutputDebugStringW(L"SerialScannerLangExt::IsPropReadable\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::IsPropWritable(_In_ long lPropNum, _Inout_ BOOL* pboolPropWrite) noexcept
{
    UNREFERENCED_PARAMETER(lPropNum);
    UNREFERENCED_PARAMETER(pboolPropWrite);
    OutputDebugStringW(L"SerialScannerLangExt::IsPropWritable\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetNMethods(_Inout_ long* plMethods) noexcept
{
    OutputDebugStringW(L"SerialScannerLangExt::GetNMethods\r\n");
    if (plMethods != NULL)
    {
        *plMethods = NUM_METHODS;
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::FindMethod(BSTR bstrMethodName, _Inout_ long* plMethodNum) noexcept
{
    WCHAR   text[256];

    wsprintfW(text, L"SerialScannerLangExt::FindMethod %s %p\r\n", bstrMethodName, plMethodNum);
    OutputDebugStringW(text);

    if ((bstrMethodName == NULL) || (plMethodNum == NULL))
        return S_FALSE;

    for (int i = 0; i < NUM_METHODS; ++i)
    {
        if (wcscmp(SerialScannerMethodNames[i], bstrMethodName) == 0)
        {
            *plMethodNum = i;
            return S_OK;
        }
    }
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetMethodName(_In_ long lMethodNum, _In_ long lMethodAlias, _Inout_ BSTR* pbstrMethodName) noexcept
{
    WCHAR   text[256];
    wsprintfW(text, L"SerialScannerLangExt::GetMethodName %i, %i\r\n", lMethodNum, lMethodAlias);
    OutputDebugStringW(text);

    if ((pbstrMethodName != NULL) && (lMethodNum < NUM_METHODS))
    {
        *pbstrMethodName = (BSTR)SerialScannerMethodNames[lMethodNum];
        return S_OK;
    }
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetNParams(_In_ long lMethodNum, _Inout_ long* plParams) noexcept
{
    OutputDebugStringW(L"SerialScannerLangExt::GetNParams\r\n");

    if (plParams == NULL)
        return S_FALSE;

    *plParams = 0;
    if (lMethodNum < NUM_METHODS)
    {
        *plParams = SerialScannerMethodParamsCount[lMethodNum];
        return S_OK;
    }

    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::GetParamDefValue(_In_ long lMethodNum, _In_ long lParamNum, _Inout_ VARIANT* pvarParamDefValue) noexcept
{
    UNREFERENCED_PARAMETER(lMethodNum);
    UNREFERENCED_PARAMETER(lParamNum);
    UNREFERENCED_PARAMETER(pvarParamDefValue);
    OutputDebugStringW(L"SerialScannerLangExt::GetParamDefValue\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::HasRetVal(_In_ long lMethodNum, _Inout_ BOOL* pboolRetValue) noexcept
{
    UNREFERENCED_PARAMETER(lMethodNum);
    OutputDebugStringW(L"SerialScannerLangExt::HasRetVal\r\n");
    
    if (pboolRetValue != NULL)
        *pboolRetValue = TRUE;
    else
        return S_FALSE;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::CallAsProc(_In_ long lMethodNum, _Inout_ SAFEARRAY** paParams) noexcept
{
    UNREFERENCED_PARAMETER(lMethodNum);
    UNREFERENCED_PARAMETER(paParams);
    OutputDebugStringW(L"SerialScannerLangExt::CallAsProc\r\n");
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE SerialScannerLangExt::CallAsFunc(_In_ long lMethodNum, _Inout_ VARIANT* pvarRetValue, _Inout_ SAFEARRAY** paParams) noexcept
{
    HRESULT     r = S_OK;
    VARIANT     fparams[5] = { 0 };

    OutputDebugStringW(L"SerialScannerLangExt::CallAsFunc\r\n");

    if ((!paParams) || (!pvarRetValue))
        return E_INVALIDARG;

    switch (lMethodNum)
    {
    case 0: // Open

        for (int c = 0; c < 5; ++c)
        {
            if (S_OK != GetNParam(*paParams, c, &fparams[c]))
            {
                r = S_FALSE;
                break;
            }
        }

        VariantInit(pvarRetValue);
        V_VT(pvarRetValue) = VT_BOOL;
        V_BOOL(pvarRetValue) = FALSE;

        if (r == S_OK)
            V_BOOL(pvarRetValue) = (VARIANT_BOOL)this->Open(
                V_BSTR(&fparams[0]),
                V_UI4(&fparams[1]),
                V_UI1(&fparams[2]),
                V_UI1(&fparams[3]),
                V_UI1(&fparams[4]));

        break;

    case 1: // Read
        VariantInit(pvarRetValue);
        V_VT(pvarRetValue) = VT_BOOL;
        V_BOOL(pvarRetValue) = (VARIANT_BOOL)this->Read(NULL);
        break;

    case 2: // Close
        VariantInit(pvarRetValue);
        V_VT(pvarRetValue) = VT_BOOL;
        V_BOOL(pvarRetValue) = (VARIANT_BOOL)this->Close();
        break;

    default:
        break;
    }

    return r;
}