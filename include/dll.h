#pragma once
#define UNICODE
#include <olectl.h>

/*************
 * DLL注册表信息
 * 这里定义了RTDServer的ProgId、CLSID、TypeId等信息
 * 这些信息用于COM注册和查找
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

/// @brief 检测当前进程是否具有写入HKLM权限 @return true:有写入权限，false:没有写入权限
bool CanWriteToHKLM();
bool CheckRegistry();
bool AutoRegisterDll();