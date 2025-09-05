#include "RtdServer.h"
#include "xllRTD.h"
#include <ks.h>

constexpr long DEFAULT_HEARTBEAT_INTERVAL = 15000; // Default heartbeat interval (milliseconds)
constexpr int DEFAULT_RUNNING_INTERVAL = 1000;     // Default running interval (milliseconds)

WCHAR RtdServer_DllPath[1024] = L"";
DWORD RtdServer::WorkerThreadProc() {
    while (m_running) {
        bool isChanged = false;

        // Use scoped lock to protect topic map
        {
            std::lock_guard<std::mutex> lock(m_TopicMapMutex);
            for (const auto& pair : m_TopicMap) {
                Topic* topic = pair.second;
                if (topic != nullptr && m_running) {
                    if (!topic->isTaskRunning()) {
                        topic->runTask();
                    }
                    if (topic->hasChanged()) {
                        isChanged = true;
                    }
                }
            }
        }

        // Call callback outside the lock to avoid deadlock
        if (isChanged && m_pCallbackObject != nullptr && m_running) {
            m_pCallbackObject->UpdateNotify();
        }

        if (!m_running) break;
        Sleep(runing_ms);
    }
    return 0;
}

RtdServer::RtdServer() : m_HeartbeatInterval(DEFAULT_HEARTBEAT_INTERVAL),
runing_ms(DEFAULT_RUNNING_INTERVAL) {
    LoadTypeInfo(&m_pTypeInfoInterface, IID_IRtdServer, 0x0);
}

RtdServer::~RtdServer() {
    // Ensure server termination and resource cleanup
    ServerTerminate();

    // Release type information interface
    if (m_pTypeInfoInterface != nullptr) {
        m_pTypeInfoInterface->Release();
        m_pTypeInfoInterface = nullptr;
    }
}

HRESULT STDMETHODCALLTYPE RtdServer::QueryInterface(REFIID riid, void** ppvObject) {
    if (IID_IUnknown == riid) {
        *ppvObject = static_cast<IUnknown*>(this);
    } else if (IID_IDispatch == riid) {
        *ppvObject = static_cast<IDispatch*>(this);
    } else if (IID_IRtdServer == riid) {
        *ppvObject = static_cast<IRtdServer*>(this);
    } else {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }
    static_cast<IUnknown*>(*ppvObject)->AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE RtdServer::AddRef(void) {
    m_RefCount++;
    return m_RefCount;
}

ULONG STDMETHODCALLTYPE RtdServer::Release(void) {
    m_RefCount--;
    if (0 == m_RefCount) {
        delete this;
        return 0;
    }
    return m_RefCount;
}

HRESULT STDMETHODCALLTYPE RtdServer::GetTypeInfoCount(UINT* pctinfo) {
    *pctinfo = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RtdServer::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE RtdServer::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) {
    HRESULT hr = E_FAIL;

    if (riid != IID_NULL)
        return E_INVALIDARG;

    if (m_pTypeInfoInterface != nullptr)
        hr = m_pTypeInfoInterface->GetIDsOfNames(rgszNames, cNames, rgDispId);

    return hr;
}

HRESULT STDMETHODCALLTYPE RtdServer::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) {
    HRESULT hr = DISP_E_PARAMNOTFOUND;

    if (riid != IID_NULL)
        return E_INVALIDARG;

    if (m_pTypeInfoInterface != nullptr) {
        hr = m_pTypeInfoInterface->Invoke(static_cast<IRtdServer*>(this), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
    }

    return hr;
}

HRESULT STDMETHODCALLTYPE RtdServer::ServerStart(IRTDUpdateEvent* CallbackObject, long* pfRes) {
    if (CallbackObject == nullptr || pfRes == nullptr) {
        return E_POINTER;
    }

    HRESULT hr = S_OK;

    // Set heartbeat interval
    hr = CallbackObject->put_HeartbeatInterval(m_HeartbeatInterval);
    if (FAILED(hr)) {
        return hr;
    }

    m_pCallbackObject = CallbackObject;

    if (!m_running) {
        m_running = true;

        // Create worker thread
        m_hThread = CreateThread(nullptr, 0, [](LPVOID param) -> DWORD {
            RtdServer* self = static_cast<RtdServer*>(param);
            return self->WorkerThreadProc(); }, this, 0, &m_threadID);

        if (m_hThread == nullptr) {
            m_running = false;
            m_threadID = 0;
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    *pfRes = static_cast<long>(m_threadID);
    return hr;
}

HRESULT STDMETHODCALLTYPE RtdServer::ConnectData(long TopicID, SAFEARRAY** Strings, VARIANT_BOOL* GetNewValues, VARIANT* pvarOut) {
    if (pvarOut == nullptr || Strings == nullptr || GetNewValues == nullptr) {
        return E_POINTER;
    }

    std::lock_guard<std::mutex> lock(m_TopicMapMutex);

    // Check if topic ID already exists
    if (m_TopicMap.find(TopicID) != m_TopicMap.end()) {
        return E_FAIL; // Topic already exists
    }

    // Create new topic
    Topic* pTopic = nullptr;
    try {
        pTopic = new Topic(TopicID, Strings, L"Default Value");

        // createRtdTask(pTopic);
        registerRTDTask(pTopic);

        // If need to get new values and has default value, return default value
        if (*GetNewValues != VARIANT_FALSE && pTopic->hasDefaultValue()) {
            VARIANT temp = createVariant(pTopic->getDefaultValue());
            *pvarOut = temp;
            // Note: Don't call VariantClear here because the value has been transferred to pvarOut
        } else {
            VariantInit(pvarOut); // Initialize to empty value
        }

        m_TopicMap[TopicID] = pTopic;
        return S_OK;
    } catch (const std::exception&) {
        delete pTopic; // Clean up partially created resources
        return E_OUTOFMEMORY;
    }
}

HRESULT STDMETHODCALLTYPE RtdServer::RefreshData(long* TopicCount, SAFEARRAY** parrayOut) {
    if (TopicCount == nullptr || parrayOut == nullptr) {
        return E_POINTER;
    }

    if (*parrayOut != nullptr) {
        return E_INVALIDARG; // Output array should be empty
    }

    if (!m_running) {
        *TopicCount = 0;
        return S_OK;
    }

    std::vector<std::pair<long, Topic*>> changedTopics;

    // Collect all changed topics
    {
        std::lock_guard<std::mutex> lock(m_TopicMapMutex);
        for (const auto& pair : m_TopicMap) {
            Topic* topic = pair.second;
            if (topic != nullptr && topic->hasChanged()) {
                changedTopics.emplace_back(pair.first, topic);
            }
        }
    }

    *TopicCount = static_cast<long>(changedTopics.size());

    if (*TopicCount == 0) {
        return S_OK; // No changed data
    }

    // Create 2D array: first dimension is 2 (TopicID and Value), second dimension is topic count
    SAFEARRAYBOUND bounds[2];
    bounds[0].cElements = 2;
    bounds[0].lLbound = 0;
    bounds[1].cElements = *TopicCount;
    bounds[1].lLbound = 0;

    *parrayOut = SafeArrayCreate(VT_VARIANT, 2, bounds);
    if (*parrayOut == nullptr) {
        return E_OUTOFMEMORY;
    }

    // Fill data
    for (long i = 0; i < *TopicCount; ++i) {
        changedTopics[i].second->update(parrayOut, i);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE RtdServer::DisconnectData(long TopicID) {
    std::lock_guard<std::mutex> lock(m_TopicMapMutex);

    auto it = m_TopicMap.find(TopicID);
    if (it != m_TopicMap.end()) {
        Topic* topic = it->second;
        if (topic != nullptr) {
            topic->stopTask(); // Stop task before deletion
            delete topic;
        }
        m_TopicMap.erase(it);
        return S_OK;
    }
    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE RtdServer::Heartbeat(long* pfRes) {
    if (pfRes == nullptr) {
        return E_POINTER;
    }
    *pfRes = m_threadID;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RtdServer::ServerTerminate() {
    // Stop running flag
    m_running = false;

    // Wait for thread termination (WaitForSingleObject will block UI, so force termination directly)
    if (m_hThread != nullptr) {
        TerminateThread(m_hThread, 0);
        CloseHandle(m_hThread);
        m_hThread = nullptr;
        m_threadID = 0;
    }

    // Clean up all topics
    {
        std::lock_guard<std::mutex> lock(m_TopicMapMutex);
        for (const auto& pair : m_TopicMap) {
            if (pair.second != nullptr) {
                pair.second->stopTask(); // Stop task
                delete pair.second;
            }
        }
        m_TopicMap.clear();
        m_DeleteTopicIDs.clear();
    }

    // Clean up callback object reference
    m_pCallbackObject = nullptr;

    return S_OK;
}

STDMETHODIMP RtdServer::LoadTypeInfo(ITypeInfo** pptinfo, REFCLSID clsid, LCID lcid) {
    HRESULT hr;
    LPTYPELIB ptlib = NULL;
    LPTYPEINFO ptinfo = NULL;
    *pptinfo = NULL;

    hr = LoadRegTypeLib(LIBID_RTDServerLib, 1, 0, lcid, &ptlib);
    if (FAILED(hr)) {
        hr = LoadTypeLib(L"EXCEL.EXE", &ptlib);
        if (FAILED(hr)) {
            hr = LoadTypeLib(L"etapi.dll", &ptlib);
            if (FAILED(hr)) {
                return hr;
            }
        }
    }
    hr = ptlib->GetTypeInfoOfGuid(clsid, &ptinfo);
    if (FAILED(hr)) {
        ptlib->Release();
        return hr;
    }
    ptlib->Release();
    *pptinfo = ptinfo;
    return S_OK;
}
