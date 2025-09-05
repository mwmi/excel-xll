#pragma once
#include "IRTDServer.h"
#include "RTDTopic.h"
#include <map>

/**
 * @brief 为RTD主题创建任务
 * @param topic 指向Topic对象的指针
 * @return 成功返回0，失败返回非0值
 */
int createRtdTask(Topic* topic);

/// DLL完整路径存储缓冲区
extern WCHAR RtdServer_DllPath[1024];

/// RTD服务器程序标识符
#define RtdServer_ProgId L"rtdserver"

/// RTD服务器类标识符字符串
#define RtdServer_CLSID L"{EC0E6192-DB51-11D3-8F3E-00C04F3651B8}"

/// 定义RTD服务器的GUID
DEFINE_GUID(CLSID_RtdServer, 0xEC0E6192, 0xDB51, 0x11D3, 0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8);

/**
 * @brief RTD服务器类，实现IRtdServer接口
 *
 * 这个类提供了RTD（Real-Time Data）服务器的功能，允许Excel等应用程序
 * 实时获取数据更新。它实现了标准的COM接口IRtdServer。
 */
class RtdServer : public IRtdServer {
private:
  /// COM对象引用计数
  ULONG m_RefCount = 0;

  /// 类型信息接口指针
  ITypeInfo* m_pTypeInfoInterface = nullptr;

  /// RTD更新事件回调对象指针
  IRTDUpdateEvent* m_pCallbackObject = nullptr;

  /// 心跳间隔时间（毫秒）
  long m_HeartbeatInterval = 15000;

  /// 主题映射表，存储主题ID与主题对象的对应关系
  std::map<long, Topic*> m_TopicMap;

  /// 保护主题映射表的互斥锁
  mutable std::mutex m_TopicMapMutex;

  /// 待删除的主题ID列表
  std::vector<long> m_DeleteTopicIDs;

  /// 工作线程句柄
  HANDLE m_hThread = nullptr;

  /// 线程ID
  DWORD m_threadID = 0;

  /// 服务器运行状态标志
  bool m_running = false;

  /// 运行间隔时间（毫秒）
  int runing_ms = 1000;

  /**
   * @brief 工作线程处理程序
   * @return DWORD 线程退出码
   */
  DWORD WorkerThreadProc();

public:
  /**
   * @brief 构造函数
   */
  RtdServer();

  /**
   * @brief 析构函数
   */
  ~RtdServer();

  // IUnknown methods
  /**
   * @brief 查询接口
   * @param riid 请求的接口ID
   * @param ppvObject 输出接口指针
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject);

  /**
   * @brief 增加引用计数
   * @return 新的引用计数值
   */
  ULONG STDMETHODCALLTYPE AddRef(void);

  /**
   * @brief 减少引用计数
   * @return 新的引用计数值
   */
  ULONG STDMETHODCALLTYPE Release(void);

  // IDispatch methods
  /**
   * @brief 获取类型信息数量
   * @param pctinfo 输出类型信息数量
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);

  /**
   * @brief 获取类型信息
   * @param iTInfo 类型信息索引
   * @param lcid 区域设置标识符
   * @param ppTInfo 输出类型信息指针
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);

  /**
   * @brief 获取方法或属性的分发ID
   * @param riid 保留参数，必须为IID_NULL
   * @param rgszNames 方法或属性名称数组
   * @param cNames 名称数量
   * @param lcid 区域设置标识符
   * @param rgDispId 输出分发ID数组
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);

  /**
   * @brief 调用方法或访问属性
   * @param dispIdMember 分发ID
   * @param riid 保留参数，必须为IID_NULL
   * @param lcid 区域设置标识符
   * @param wFlags 调用标志
   * @param pDispParams 参数信息
   * @param pVarResult 返回值
   * @param pExcepInfo 异常信息
   * @param puArgErr 错误参数索引
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

  // IRtdServer methods
  /**
   * @brief 启动RTD服务器
   * @param CallbackObject 回调对象指针
   * @param pfRes 输出结果标志
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE ServerStart(IRTDUpdateEvent* CallbackObject, long* pfRes);

  /**
   * @brief 连接数据主题
   * @param TopicID 主题ID
   * @param Strings 主题参数字符串数组
   * @param GetNewValues 是否获取新值的标志
   * @param pvarOut 输出数据
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE ConnectData(long TopicID, SAFEARRAY** Strings, VARIANT_BOOL* GetNewValues, VARIANT* pvarOut);

  /**
   * @brief 刷新数据
   * @param TopicCount 输出主题数量
   * @param parrayOut 输出数据数组
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE RefreshData(long* TopicCount, SAFEARRAY** parrayOut);

  /**
   * @brief 断开数据主题连接
   * @param TopicID 主题ID
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE DisconnectData(long TopicID);

  /**
   * @brief 心跳检测
   * @param pfRes 输出结果标志
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE Heartbeat(long* pfRes);

  /**
   * @brief 终止RTD服务器
   * @return HRESULT 操作结果
   */
  HRESULT STDMETHODCALLTYPE ServerTerminate();

  // User-defined methods
  /**
   * @brief 加载类型信息的辅助方法
   * @param pptinfo 输出类型信息指针
   * @param clsid 类标识符
   * @param lcid 区域设置标识符
   * @return HRESULT 操作结果
   */
  STDMETHODIMP LoadTypeInfo(ITypeInfo** pptinfo, REFCLSID clsid, LCID lcid);
};
