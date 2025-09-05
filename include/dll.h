#pragma once
#define UNICODE
#include <olectl.h>

/*************
 * DLL registry information
 * Here defines RTDServer's ProgId, CLSID, TypeId and other information
 * This information is used for COM registration and lookup
 *************/
extern const WCHAR* regTable[][3];

class CComFactory : public IClassFactory {
private:
    ULONG m_RefCount = 0;

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(const IID& riid, void** ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppvObject);
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock);
};

STDAPI DllRegisterServer();
STDAPI DllUnregisterServer();
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv);
STDAPI DllCanUnloadNow();

/// @brief Check if current process has write permission to HKLM @return true:has write permission, false:no write permission
bool CanWriteToHKLM();
bool CheckRegistry();
bool AutoRegisterDll();