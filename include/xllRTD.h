/**
 * @file xllRTD.h
 * @brief RTD (Real-Time Data) function registration and call management module
 * @author ExcelXLL Framework
 * @date 2025
 *
 * This file defines the registration, management and call mechanism of RTD functions, providing a unified interface
 * for handling real-time data functions in Excel. The RTD system allows Excel worksheets to obtain
 * and update data from external data sources in real time.
 *
 * Main functions:
 * - RTD function registration and management
 * - Asynchronous function call support
 * - Default value and state management
 * - Excel RTD API encapsulation
 *
 * @see RtdServer.h RTD server implementation
 * @see RTDTopic.h RTD topic management
 */
#pragma once
#include "xllType.h"
#include "RtdServer.h"

 /// @brief RTD function pointer type definition, used to define asynchronous data acquisition functions
 /// @param args Function parameter list, using smart pointers to manage xllType objects
 /// @param topic RTD topic object pointer, containing topic information and state
 /// @return Execution result, 0 for success, negative for error
using RtdFun = int (*)(xllptrlist, Topic*);

/**
 * @struct RTDRegister
 * @brief RTD function registration and management class, using singleton pattern
 *
 * This class is responsible for managing all RTD functions, including registration, calling and state management.
 * Using singleton pattern ensures only one registry instance exists globally.
 *
 * Design features:
 * - Using singleton pattern to ensure global uniqueness
 * - Support function registration, lookup and calling
 * - Manage function default values and asynchronous status
 * - Thread-safe function registry
 *
 * Usage example:
 * ```cpp
 * RTDRegister& rtd = RTDRegister::instance();
 * rtd.registerRTDFunction(L"myFunc", myFuncPtr, L"Loading...", true);
 * ```
 */
struct RTDRegister {
private:
    /// @brief RTD function mapping table, storing mapping relationship from function names to function pointers
    std::map<std::wstring, RtdFun> _async_functions;
    /// @brief Function default value mapping table, storing initial return values for each function
    std::map<std::wstring, std::wstring> _default_values;
    /// @brief Function asynchronous status mapping table, indicating whether functions execute in separate threads
    std::map<std::wstring, bool> _is_async;
public:
    /**
     * @brief Get singleton object of RTD register
     * @return RTDRegister& Globally unique RTD register instance
     * @note Uses static local variable to implement thread-safe singleton pattern
     */
    static RTDRegister& instance();

    /**
     * @brief Register RTD function, supporting default values and async flags
     * @param name Function name, used as first parameter in RTD expression
     * @param fun Function pointer, pointing to actual data acquisition function
     * @param default_value Default return value, displayed during function execution
     * @param is_async Whether to execute asynchronously, default false
     */
    void registerRTDFunction(const std::wstring& name, RtdFun fun, const wchar_t* default_value = L"", bool is_async = false);

    /**
     * @brief Register RTD function (simplified version)
     * @param name Function name
     * @param fun Function pointer
     * @param is_async Whether to execute asynchronously
     */
    void registerRTDFunction(const std::wstring& name, RtdFun fun, bool is_async);

    /**
     * @brief Run RTD function
     * @param name Function name
     * @param args Function parameter list (move semantics)
     * @param topic RTD topic object pointer
     * @return Execution result, 0 for success, -1 for failure (function not registered)
     */
    int runAsyncFunction(const std::wstring& name, xllptrlist& args, Topic* topic);

    /// @brief Check if function is registered @param name Function name @return Whether registered
    bool isFunctionRegistered(const std::wstring& name);

    /**
     * @brief Get function's default value
     * @param name Function name
     * @param default_value Output parameter, stores default value
     * @return Returns true on success, false on failure
     */
    bool getDefaultValue(const std::wstring& name, std::wstring& default_value);

    /**
     * @brief Check if function executes asynchronously
     * @param name Function name
     * @return Returns true for async, false for sync
     */

    bool isFunctionAsync(const std::wstring& name);

    /**
     * @brief Get function pointer
     * @param name Function name
     * @return Function pointer, returns nullptr if not exists
     * @warning Should use isFunctionRegistered to check function existence before calling
     */
    RtdFun& getFunction(const std::wstring& name);
};

/**
 * @brief Excel RTD function template wrapper for calling Excel's RTD API
 *
 * This function wraps parameters into Excel-recognizable format and calls Excel's
 * xlfRtd function to implement real-time data acquisition.
 *
 * @tparam Args Variadic parameter type list
 * @param result Output parameter, stores Excel's return result
 * @param args RTD function parameter list
 * @return Excel API call result, xlretSuccess indicates success
 *
 * @note First parameter will be used as RTD server function name
 * @see Excel12v Low-level Excel API function
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
        result = L"RTD service exception";
    }
    return ret;
}

/**
 * @brief RTD task registration function
 *
 * Based on Topic's parameter information, finds corresponding function in RTD registry
 * and sets appropriate task and default values for the Topic.
 *
 * @param topic RTD topic object pointer, contains parameter information
 * @return Registration result, 0 for success, negative for failure
 *
 * @note First parameter is treated as function name, remaining parameters are passed to registered function
 * @see RTDRegister::runAsyncFunction
 */
int registerRTDTask(Topic* topic);

