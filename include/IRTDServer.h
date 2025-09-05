#pragma once
#include <windows.h>

DEFINE_GUID(LIBID_RTDServerLib, 0x00020813, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_IRtdServer, 0xec0e6191, 0xdb51, 0x11d3, 0x8f, 0x3e, 0x00, 0xc0, 0x4f, 0x36, 0x51, 0xb8);
struct IRtdServer : IDispatch {
    virtual HRESULT __stdcall ServerStart(
        /*[in]*/ struct IRTDUpdateEvent* CallbackObject,
        /*[out,retval]*/ long* pfRes) = 0;
    virtual HRESULT __stdcall ConnectData(
        /*[in]*/ long TopicID,
        /*[in]*/ SAFEARRAY** Strings,
        /*[in,out]*/ VARIANT_BOOL* GetNewValues,
        /*[out,retval]*/ VARIANT* pvarOut) = 0;
    virtual HRESULT __stdcall RefreshData(
        /*[in,out]*/ long* TopicCount,
        /*[out,retval]*/ SAFEARRAY** parrayOut) = 0;
    virtual HRESULT __stdcall DisconnectData(
        /*[in]*/ long TopicID) = 0;
    virtual HRESULT __stdcall Heartbeat(
        /*[out,retval]*/ long* pfRes) = 0;
    virtual HRESULT __stdcall ServerTerminate() = 0;
};

DEFINE_GUID(IID_IRTDUpdateEvent, 0xa43788c1, 0xd91b, 0x11d3, 0x8f, 0x39, 0x00, 0xc0, 0x4f, 0x36, 0x51, 0xb8);
struct IRTDUpdateEvent : IDispatch {
    virtual HRESULT __stdcall UpdateNotify() = 0;
    virtual HRESULT __stdcall get_HeartbeatInterval(
        /*[out,retval]*/ long* plRetVal) = 0;
    virtual HRESULT __stdcall put_HeartbeatInterval(
        /*[in]*/ long plRetVal) = 0;
    virtual HRESULT __stdcall Disconnect() = 0;
};