#pragma once
#include "IRTDServer.h"
#include "RTDTopic.h"
#include <map>

/**
 * @brief Create task for RTD topic
 * @param topic Pointer to Topic object
 * @return Returns 0 on success, non-zero on failure
 */
int createRtdTask(Topic* topic);

/// DLL full path storage buffer
extern WCHAR RtdServer_DllPath[1024];

/// RTD server program identifier
#define RtdServer_ProgId L"rtdserver"

/// RTD server class identifier string
#define RtdServer_CLSID L"{EC0E6192-DB51-11D3-8F3E-00C04F3651B8}"

/// Define GUID for RTD server
DEFINE_GUID(CLSID_RtdServer, 0xEC0E6192, 0xDB51, 0x11D3, 0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8);

/**
 * @brief RTD server class, implements IRtdServer interface
 *
 * This class provides RTD (Real-Time Data) server functionality, allowing applications like Excel
 * to obtain real-time data updates. It implements the standard COM interface IRtdServer.
 */
class RtdServer : public IRtdServer {
private:
  /// COM object reference count
  ULONG m_RefCount = 0;

  /// Type information interface pointer
  ITypeInfo* m_pTypeInfoInterface = nullptr;

  /// RTD update event callback object pointer
  IRTDUpdateEvent* m_pCallbackObject = nullptr;

  /// Heartbeat interval time (milliseconds)
  long m_HeartbeatInterval = 15000;

  /// Topic map, stores correspondence between topic IDs and topic objects
  std::map<long, Topic*> m_TopicMap;

  /// Mutex protecting the topic map
  mutable std::mutex m_TopicMapMutex;

  /// List of topic IDs to be deleted
  std::vector<long> m_DeleteTopicIDs;

  /// Worker thread handle
  HANDLE m_hThread = nullptr;

  /// Thread ID
  DWORD m_threadID = 0;

  /// Server running status flag
  bool m_running = false;

  /// Running interval time (milliseconds)
  int runing_ms = 1000;

  /**
   * @brief Worker thread procedure
   * @return DWORD Thread exit code
   */
  DWORD WorkerThreadProc();

public:
  /**
   * @brief Constructor
   */
  RtdServer();

  /**
   * @brief Destructor
   */
  ~RtdServer();

  // IUnknown methods
  /**
   * @brief Query interface
   * @param riid Requested interface ID
   * @param ppvObject Output interface pointer
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject);

  /**
   * @brief Increase reference count
   * @return New reference count value
   */
  ULONG STDMETHODCALLTYPE AddRef(void);

  /**
   * @brief Decrease reference count
   * @return New reference count value
   */
  ULONG STDMETHODCALLTYPE Release(void);

  // IDispatch methods
  /**
   * @brief Get type information count
   * @param pctinfo Output type information count
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);

  /**
   * @brief Get type information
   * @param iTInfo Type information index
   * @param lcid Locale identifier
   * @param ppTInfo Output type information pointer
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);

  /**
   * @brief Get dispatch ID of method or property
   * @param riid Reserved parameter, must be IID_NULL
   * @param rgszNames Method or property name array
   * @param cNames Number of names
   * @param lcid Locale identifier
   * @param rgDispId Output dispatch ID array
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);

  /**
   * @brief Call method or access property
   * @param dispIdMember Dispatch ID
   * @param riid Reserved parameter, must be IID_NULL
   * @param lcid Locale identifier
   * @param wFlags Call flags
   * @param pDispParams Parameter information
   * @param pVarResult Return value
   * @param pExcepInfo Exception information
   * @param puArgErr Error parameter index
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

  // IRtdServer methods
  /**
   * @brief Start RTD server
   * @param CallbackObject Callback object pointer
   * @param pfRes Output result flag
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE ServerStart(IRTDUpdateEvent* CallbackObject, long* pfRes);

  /**
   * @brief Connect data topic
   * @param TopicID Topic ID
   * @param Strings Topic parameter string array
   * @param GetNewValues Flag to get new values
   * @param pvarOut Output data
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE ConnectData(long TopicID, SAFEARRAY** Strings, VARIANT_BOOL* GetNewValues, VARIANT* pvarOut);

  /**
   * @brief Refresh data
   * @param TopicCount Output topic count
   * @param parrayOut Output data array
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE RefreshData(long* TopicCount, SAFEARRAY** parrayOut);

  /**
   * @brief Disconnect data topic connection
   * @param TopicID Topic ID
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE DisconnectData(long TopicID);

  /**
   * @brief Heartbeat detection
   * @param pfRes Output result flag
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE Heartbeat(long* pfRes);

  /**
   * @brief Terminate RTD server
   * @return HRESULT Operation result
   */
  HRESULT STDMETHODCALLTYPE ServerTerminate();

  // User-defined methods
  /**
   * @brief Helper method to load type information
   * @param pptinfo Output type information pointer
   * @param clsid Class identifier
   * @param lcid Locale identifier
   * @return HRESULT Operation result
   */
  STDMETHODIMP LoadTypeInfo(ITypeInfo** pptinfo, REFCLSID clsid, LCID lcid);
};
