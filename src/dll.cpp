
#define INITGUID
#include "dll.h"
#include "RtdServer.h"
#include "xllManager.h"

const WCHAR* regTable[][3] = {
    {L"Software\\Classes\\" RtdServer_ProgId, 0, RtdServer_ProgId},
    {L"Software\\Classes\\" RtdServer_ProgId L"\\CLSID", 0, RtdServer_CLSID},
    {L"Software\\Classes\\CLSID\\" RtdServer_CLSID, 0, RtdServer_ProgId},
    {L"Software\\Classes\\CLSID\\" RtdServer_CLSID L"\\InprocServer32", 0, (const WCHAR*)-1},
    {L"Software\\Classes\\CLSID\\" RtdServer_CLSID L"\\ProgId", 0, RtdServer_ProgId},
};

BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD ul_reason_for_call,
    LPVOID lpReserved) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
    GetModuleFileName(hModule, RtdServer_DllPath, sizeof(RtdServer_DllPath) / sizeof(WCHAR));
    xll::init::instance().call();
    break;
    case DLL_PROCESS_DETACH:
    break;
    }
    return TRUE;
}

HRESULT STDMETHODCALLTYPE CComFactory::QueryInterface(const IID& riid, void** ppv) {
    if (IID_IUnknown == riid) {
        *ppv = static_cast<IUnknown*>(this);
    } else if (IID_IClassFactory == riid) {
        *ppv = static_cast<IClassFactory*>(this);
    } else {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    static_cast<IUnknown*>(*ppv)->AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE CComFactory::AddRef() {
    m_RefCount++;
    return m_RefCount;
}

ULONG STDMETHODCALLTYPE CComFactory::Release() {
    m_RefCount--;
    if (0 == m_RefCount) {
        delete this;
        return 0;
    }
    return m_RefCount;
}

HRESULT STDMETHODCALLTYPE CComFactory::CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppvObject) {
    if (NULL != pUnkOuter) {
        return CLASS_E_NOAGGREGATION;
    }
    HRESULT hr = E_OUTOFMEMORY;
    RtdServer* pObj = new RtdServer();
    if (NULL == pObj) {
        return hr;
    }
    hr = pObj->QueryInterface(riid, ppvObject);
    if (S_OK != hr) {
        delete pObj;
    }
    return hr;
}

HRESULT STDMETHODCALLTYPE CComFactory::LockServer(BOOL fLock) {
    return NOERROR;
}

STDAPI DllRegisterServer() {
    HKEY hk = HKEY_CURRENT_USER;
    if (CanWriteToHKLM()) {
        hk = HKEY_LOCAL_MACHINE;
    }

    if (RtdServer_DllPath[0] == L'\0') return S_FALSE;

    int n = sizeof(regTable) / sizeof(regTable[0]);
    for (int i = 0; i < n; i++) {
        const WCHAR* key = regTable[i][0];
        const WCHAR* valueName = regTable[i][1];
        const WCHAR* value = regTable[i][2];

        if (value == (const WCHAR*)-1) {
            value = RtdServer_DllPath;
        }

        HKEY hkey;
        long err = RegCreateKey(hk, key, &hkey);
        if (err == ERROR_SUCCESS) {
            err = RegSetValueEx(hkey, valueName, 0, REG_SZ, (const BYTE*)value, (lstrlen(value) + 1) * sizeof(WCHAR));
            RegCloseKey(hkey);
        }

        if (err != ERROR_SUCCESS) {
            DllUnregisterServer();
            return SELFREG_E_CLASS;
        }
    }
    return S_OK;
}

STDAPI DllUnregisterServer() {
    HRESULT hrm = S_OK, hru = S_OK;
    int n = sizeof(regTable) / sizeof(regTable[0]);
    if (CanWriteToHKLM()) {
        for (int i = n - 1; i >= 0; i--) {
            const WCHAR* key = regTable[i][0];
            long err = RegDeleteKey(HKEY_LOCAL_MACHINE, key);
            if (err != ERROR_SUCCESS) {
                hrm = S_FALSE;
            }
        }
    } else {
        hrm = S_FALSE;
    }
    for (int i = n - 1; i >= 0; i--) {
        const WCHAR* key = regTable[i][0];
        long err = RegDeleteKey(HKEY_CURRENT_USER, key);
        if (err != ERROR_SUCCESS) {
            hru = S_FALSE;
        }
    }
    return hrm == S_OK || hru == S_OK ? S_OK : SELFREG_E_CLASS;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
    if (CLSID_RtdServer != rclsid) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    CComFactory* pFactory = new CComFactory();
    if (NULL == pFactory) {
        return E_OUTOFMEMORY;
    }
    HRESULT result = pFactory->QueryInterface(riid, ppv);
    return result;
}

STDAPI DllCanUnloadNow() {
    return S_OK;
}

bool CanWriteToHKLM() {
    HKEY hKey = nullptr;
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Software\\Classes", 0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return true;
    }
    return false;
}

bool CheckRegistry() {
    bool ret = false;
    HKEY hKey = nullptr;
    WCHAR path[1024] = {0};
    DWORD dwType = REG_SZ;
    DWORD dwSize = sizeof(path);
    const WCHAR* subKey = L"Software\\Classes\\CLSID\\" RtdServer_CLSID L"\\InprocServer32";

    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, subKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        if (RegQueryValueEx(hKey, nullptr, nullptr, &dwType, reinterpret_cast<LPBYTE>(path), &dwSize) == ERROR_SUCCESS) {
            ret = (wcscmp(path, RtdServer_DllPath) == 0);
        }
        RegCloseKey(hKey);
        if (ret) return true;
    }

    if (CanWriteToHKLM()) return false;

    if (RegOpenKeyEx(HKEY_CURRENT_USER, subKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        if (RegQueryValueEx(hKey, nullptr, nullptr, &dwType, reinterpret_cast<LPBYTE>(path), &dwSize) == ERROR_SUCCESS) {
            ret = (wcscmp(path, RtdServer_DllPath) == 0);
        }
        RegCloseKey(hKey);
        if (ret) return true;
    }

    return false;
}

bool AutoRegisterDll() {
    if (!CheckRegistry()) {
        return DllRegisterServer() == S_OK;
    }
    return true;
}