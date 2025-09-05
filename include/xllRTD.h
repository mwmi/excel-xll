/**
 * @file xllRTD.h
 * @brief RTD（实时数据）函数注册和调用管理模块
 * @author ExcelXLL Framework
 * @date 2025
 *
 * 本文件定义了RTD函数的注册、管理和调用机制，提供了一个统一的接口
 * 用于处理Excel中的实时数据函数。RTD系统允许Excel工作表实时获取
 * 和更新外部数据源的数据。
 *
 * 主要功能：
 * - RTD函数注册和管理
 * - 异步函数调用支持
 * - 默认值和状态管理
 * - Excel RTD API封装
 *
 * @see RtdServer.h RTD服务器实现
 * @see RTDTopic.h RTD主题管理
 */
#pragma once
#include "xllType.h"
#include "RtdServer.h"

 /// @brief RTD函数指针类型定义，用于定义异步数据获取函数
 /// @param args 函数参数列表，使用智能指针管理xllType对象
 /// @param topic RTD主题对象指针，包含主题信息和状态
 /// @return 执行结果，0表示成功，负值表示错误
using RtdFun = int (*)(xllptrlist, Topic*);

/**
 * @struct RTDRegister
 * @brief RTD函数注册和管理类，采用单例模式
 *
 * 该类负责管理所有的RTD函数，包括注册、调用和状态管理。
 * 使用单例模式确保在全局范围内只有一个注册表实例。
 *
 * 设计特点：
 * - 采用单例模式，确保全局唯一性
 * - 支持函数注册、查找和调用
 * - 管理函数的默认值和异步状态
 * - 线程安全的函数注册表
 *
 * 使用示例：
 * ```cpp
 * RTDRegister& rtd = RTDRegister::instance();
 * rtd.registerRTDFunction(L"myFunc", myFuncPtr, L"加载中...", true);
 * ```
 */
struct RTDRegister {
private:
    /// @brief RTD函数映射表，存储函数名到函数指针的映射关系
    std::map<std::wstring, RtdFun> _async_functions;
    /// @brief 函数默认值映射表，存储每个函数的初始返回值
    std::map<std::wstring, std::wstring> _default_values;
    /// @brief 函数异步状态映射表，标识函数是否在单独线程中执行
    std::map<std::wstring, bool> _is_async;
public:
    /**
     * @brief 获取RTD注册器的单例对象
     * @return RTDRegister& 全局唯一的RTD注册器实例
     * @note 使用静态局部变量实现线程安全的单例模式
     */
    static RTDRegister& instance();

    /**
     * @brief 注册RTD函数，支持默认值和异步标志
     * @param name 函数名称，在RTD表达式中作为第一个参数使用
     * @param fun 函数指针，指向实际的数据获取函数
     * @param default_value 默认返回值，在函数执行期间显示
     * @param is_async 是否异步执行，默认false
     */
    void registerRTDFunction(const std::wstring& name, RtdFun fun, const wchar_t* default_value = L"", bool is_async = false);

    /**
     * @brief 注册RTD函数（简化版本）
     * @param name 函数名称
     * @param fun 函数指针
     * @param is_async 是否异步执行
     */
    void registerRTDFunction(const std::wstring& name, RtdFun fun, bool is_async);

    /**
     * @brief 运行RTD函数
     * @param name 函数名称
     * @param args 函数参数列表（移动语义）
     * @param topic RTD主题对象指针
     * @return 执行结果，0成功，-1失败（函数未注册）
     */
    int runAsyncFunction(const std::wstring& name, xllptrlist& args, Topic* topic);

    /// @brief 判断函数是否已注册 @param name 函数名 @return 是否已注册
    bool isFunctionRegistered(const std::wstring& name);

    /**
     * @brief 获取函数的默认值
     * @param name 函数名称
     * @param default_value 输出参数，存储默认值
     * @return 成功返回true，失败返回false
     */
    bool getDefaultValue(const std::wstring& name, std::wstring& default_value);

    /**
     * @brief 判断函数是否为异步执行
     * @param name 函数名称
     * @return 异步返回true，同步返回false
     */

    bool isFunctionAsync(const std::wstring& name);

    /**
     * @brief 获取函数指针
     * @param name 函数名称
     * @return 函数指针，如果不存在则返回nullptr
     * @warning 调用前应先使用 isFunctionRegistered 检查函数是否存在
     */
    RtdFun& getFunction(const std::wstring& name);
};

/**
 * @brief Excel RTD函数的模板封装，用于调用Excel的RTD API
 *
 * 该函数将参数包装成Excel能够识别的格式，并调用Excel的
 * xlfRtd函数来实现实时数据获取。
 *
 * @tparam Args 可变参数类型列表
 * @param result 输出参数，存储Excel返回的结果
 * @param args RTD函数的参数列表
 * @return Excel API调用结果，xlretSuccess表示成功
 *
 * @note 第一个参数会被作为RTD服务器的函数名使用
 * @see Excel12v 底层Excel API函数
 */
template <typename... Args>
int xllRTD(xllType& result, Args... args) {
    constexpr int n = sizeof...(args);
    int ret = xlretSuccess;
    xloper12 r;
    xloper12* v[n + 2];

    xllType arr[] = {[&]() {
        return xllType(args);
    }()...};

    xllType progId(RtdServer_ProgId);
    xllType server(L"");
    v[0] = progId.to_xloper12();
    v[1] = server.to_xloper12();
    v[2] = arr[0].to_xloper12();

    for (int i = 1; i < n; i++) {
        v[i + 2] = arr[i].serialize()->to_xloper12();
    }

    ret = Excel12v(xlfRtd, &r, n + 2, v);
    if (ret == xlretSuccess) {
        result = r;
        result.deserialize();
        Excel12(xlFree, 0, 1, &r);
    } else {
        result = L"RTD服务异常";
    }
    return ret;
}

/**
 * @brief RTD任务注册函数
 *
 * 根据Topic的参数信息，查找RTD注册表中的对应函数，
 * 并为该Topic设置适当的任务和默认值。
 *
 * @param topic RTD主题对象指针，包含参数信息
 * @return 注册结果，0表示成功，负值表示失败
 *
 * @note 第一个参数被视为函数名，其余参数传递给注册的函数
 * @see RTDRegister::runAsyncFunction
 */
int registerRTDTask(Topic* topic);

