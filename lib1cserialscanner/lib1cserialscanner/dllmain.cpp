/*
*  Project name:    lib1cserialscanner
*
*  File name:       dllmain.c
*  Created on:      Sep 25, 2021
*  Modified on:     Oct 3, 2021
* 
*/

#include <Windows.h>
#include <OleCtl.h>
#include "IClassFactory.h"

static HINSTANCE g_libInstance = NULL;

LSTATUS RegWriteStringValueW(
    HKEY hKey,
    LPCWSTR ValueName,
    LPCWSTR Value)
    noexcept
{
    return RegSetValueExW(hKey, ValueName, 0, REG_SZ, (const BYTE*)Value, (DWORD)(1 + wcslen(Value)) * sizeof(WCHAR));
}

HRESULT APIENTRY DllRegisterServer(void)
{
    WCHAR   dllname[MAX_PATH] = { 0 };
    HKEY    hkey = NULL, hrkey = NULL;
    HRESULT result = SELFREG_E_CLASS;

    while (0 < GetModuleFileNameW(g_libInstance, dllname, MAX_PATH))
    {
        WCHAR keyname[256] = { 0 };

        wcscpy_s(keyname, L"CLSID\\");
        if (0 == StringFromGUID2(CLSID_lib1CSerialScanner, keyname + wcslen(keyname), 128))
            break;

        if (ERROR_SUCCESS != RegCreateKeyExW(HKEY_CLASSES_ROOT, keyname, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hkey, NULL, L"Serial scanner extension for 1C"))
            break;
        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"InprocServer32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, dllname))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, L"ThreadingModel", L"Both"))
            break;

        RegCloseKey(hrkey);
        hrkey = NULL;

        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"ProgID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, L"LIB1CEXT.SerialScanner.1"))
            break;

        RegCloseKey(hrkey);
        hrkey = NULL;

        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"VersionIndependentProgID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, L"LIB1CEXT.SerialScanner"))
            break;

        RegCloseKey(hrkey);
        hrkey = NULL;
        RegCloseKey(hkey);
        hkey = NULL;

        if (ERROR_SUCCESS != RegCreateKeyExW(HKEY_CLASSES_ROOT, L"LIB1CEXT.SerialScanner", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hkey, NULL, L"LIB1CEXT.SerialScanner"))
            break;
        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"CLSID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;

        if (0 == StringFromGUID2(CLSID_lib1CSerialScanner, keyname, 128))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, keyname))
            break;

        RegCloseKey(hrkey);
        hrkey = NULL;

        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"CurVer", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, L"LIB1CEXT.SerialScanner.1"))
            break;

        RegCloseKey(hrkey);
        hrkey = NULL;
        RegCloseKey(hkey);
        hkey = NULL;

        if (ERROR_SUCCESS != RegCreateKeyExW(HKEY_CLASSES_ROOT, L"LIB1CEXT.SerialScanner.1", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hkey, NULL, L"LIB1CEXT.SerialScanner.1"))
            break;
        if (ERROR_SUCCESS != RegCreateKeyExW(hkey, L"CLSID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hrkey, NULL))
            break;
        if (0 == StringFromGUID2(CLSID_lib1CSerialScanner, keyname, 128))
            break;
        if (ERROR_SUCCESS != RegWriteStringValueW(hrkey, NULL, keyname))
            break;

        result = S_OK;
        break;
    }

    if (hkey != NULL)
        RegCloseKey(hkey);

    if (hrkey != NULL)
        RegCloseKey(hrkey);

    return result;
}

HRESULT APIENTRY DllUnregisterServer() noexcept
{
    HKEY    hkey = NULL;
    HRESULT result = S_OK;

    while (ERROR_SUCCESS == RegOpenKeyExW(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_ALL_ACCESS, &hkey))
    {
        WCHAR s_clsid[128] = { 0 };

        if (0 == StringFromGUID2(CLSID_lib1CSerialScanner, s_clsid, 128))
        {
            result = SELFREG_E_CLASS;
            break;
        }

        if (ERROR_SUCCESS != RegDeleteTreeW(hkey, s_clsid))
            result = SELFREG_E_CLASS;
        RegCloseKey(hkey);
        hkey = NULL;

        if (ERROR_SUCCESS == RegOpenKeyExW(HKEY_CLASSES_ROOT, NULL, 0, KEY_ALL_ACCESS, &hkey))
        {
            if (ERROR_SUCCESS != RegDeleteTreeW(hkey, L"LIB1CEXT.SerialScanner"))
                result = SELFREG_E_CLASS;
            if (ERROR_SUCCESS != RegDeleteTreeW(hkey, L"LIB1CEXT.SerialScanner.1"))
                result = SELFREG_E_CLASS;
        }
        else
            result = SELFREG_E_CLASS;

        break;
    }

    if (hkey != NULL)
        RegCloseKey(hkey);

    return result;
}

__control_entrypoint(DllExport)
HRESULT APIENTRY DllCanUnloadNow() noexcept
{
    OutputDebugStringW(L"DllCanUnloadNow\r\n");
    if (g_ServerLock == 0)
    {
        OutputDebugStringW(L"DllCanUnloadNow OK\r\n");
        return S_OK;
    }

    return S_FALSE;
}

_Check_return_
HRESULT APIENTRY DllGetClassObject(
    _In_ REFCLSID       rclsid,
    _In_ REFIID         riid,
    _Outptr_ LPVOID     *ppv)
{
    HRESULT hr = E_OUTOFMEMORY;

    OutputDebugStringW(L"DllGetClassObject 1\r\n");

    if (ppv == NULL)
        return E_INVALIDARG;

    *ppv = NULL;
    if (rclsid != CLSID_lib1CSerialScanner)
        return CLASS_E_CLASSNOTAVAILABLE;

    ScannerClassFactory *v = new ScannerClassFactory();
    if (v != NULL) {
        OutputDebugStringW(L"DllGetClassObject 2\r\n");
        hr = v->QueryInterface(riid, ppv);
    }

    return hr;    
}

BOOL WINAPI DllMain(
    HINSTANCE hinstDLL,     // handle to DLL module
    DWORD fdwReason,        // reason for calling function
    LPVOID lpReserved)      // reserved
    noexcept
{
    UNREFERENCED_PARAMETER(lpReserved);

    switch (fdwReason)
    {
    case DLL_PROCESS_ATTACH:
        InitializeCriticalSection(&g_ServerCS);
        g_libInstance = hinstDLL;
        break;

    case DLL_THREAD_ATTACH:
        break;

    case DLL_THREAD_DETACH:
        break;

    case DLL_PROCESS_DETACH:
        DeleteCriticalSection(&g_ServerCS);
        break;

    default:
        /* stub */
        ;
    }

    return TRUE;
}