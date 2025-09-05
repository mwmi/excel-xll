/**
 * @file xllmacros.h
 * @author mwm
 * @brief Simple macro function definitions
 * @version 0.2
 * @date 2025-08-30
 * @copyright Copyright (c) 2024
 */
#pragma once
 // Long text
#define __T(x) L##x
/// @brief Function header abbreviation
#define Function extern "C" __declspec(dllexport) LPXLOPER12
/// @brief Parameter name abbreviation
#define Param LPXLOPER12

/// @brief Expand parentheses @param X Function parameters
#define EXPAND(X) ESC(ISH X)
#define ISH(...) ISH { __VA_ARGS__ }
#define ESC(...) ESC_(__VA_ARGS__)
#define ESC_(...) VAN##__VA_ARGS__
#define VANISH

/**
 * @brief Define and register UDF (User Defined Function) for creating Excel custom functions
 * @param func Function name, will be used as the callable function name in Excel
 * @param desc Function description, can be a string or map-like format configuration
 *             Supported formats: `L"Simple description"` or `({key, value}, {key, value}, ...)`
 * @param ... Function parameter list, defines Excel parameters accepted by the UDF function
 *
 * __Working Principle__:
 * 1. Declare and define an external export function (callable by Excel)
 * 2. Create static registration class instance, complete function registration in constructor:
 *    - Register basic function information to UDFRegistry (name, parameter count, help, etc.)
 *    - Set function description and configuration through set_info
 * 3. Re-declare function, allowing implementation of specific function body after macro
 *
 * __Function Configuration Parameters__:
 * `desc` can be simple string description or complex configuration mapping:
 * - Simple format: `L"Function description"`
 * - Complex format:
 * ```cpp
 * (
 * { udf::name , L"value1" }, { udf::arguments, L"value2" },
 * { udf::help , L"value3" }, { udf::args_help, L"value4" },
 * { udf::category , L"value5" }, {udf::registername, L"value6" },
 * { udf::type , L"value7" }
 * )
 * ```
 * - Supported configuration items: `help`(help), `category`(category), `arguments`(parameter description), etc.
 *
 * __Calling in Excel__:
 * Use `=FunctionName(param1,param2,...)` in Excel cells to call
 *
 * @note UDF functions are especially suitable for custom calculation logic, such as:
 *       - Mathematical operation functions (addition, summation, statistics)
 *       - String processing functions (concatenation, formatting)
 *       - Data conversion and validation functions
 *       - Business logic calculation functions
 *
 * @warning Ensure function parameter types are handled correctly, use xllType class for type checking and conversion
 *
 * @see UDFRegistry UDF function registration manager
 * @see UDFCONFIG UDF function configuration macro
 * @see xllType Excel data type wrapper class
 * @see SET UDF initialization configuration macro
 *
 * __Usage Example 1: Simple addition function__
 *
 * ```cpp
 * UDF(Add, L"Add two numbers", Param a, Param b) {
 *     xllType result;
 *     xllType a_ = a;
 *     xllType b_ = b;
 *     if (a_.is_num() && b_.is_num()) {
 *         result = a_.get_num() + b_.get_num();
 *     } else {
 *         result = L"Type error!";
 *     }
 *     return result.get_return();
 * }
 * ```
 *
 * __Usage Example 2: String processing function with complex configuration__
 *
 * ```cpp
 * UDF(Concat2,
 *     ({udf::help, L"Concatenate two strings"},
 *      {udf::category, L"Text Functions"},
 *      {udf::arguments, L"Text1,Text2"}),
 *      Param a, Param b) {
 *     xllType result;
 *     xllType a_ = a;
 *     xllType b_ = b;
 *     std::wstring str = a_.get_str() + b_.get_str();
 *     result = str;
 *     return result.get_return();
 * }
 * ```
 *
 * __Usage Example 3: Array processing function__
 *
 * ```cpp
 * UDF(MySum, L"Array summation", Param a) {
 *     xllType result;
 *     xllType a_ = a;
 *     double sum = 0;
 *     if (a_.is_array()) {
 *         for (auto i : a_) {
 *             if (i->is_num())
 *                 sum += i->get_num();
 *         }
 *     }
 *     result = sum;
 *     return result.get_return();
 * }
 * ```
 *
 * __Usage Example 4: Function returning array__
 *
 * ```cpp
 * UDF(RetArray, L"Return test array") {
 *     xllType result;
 *     xllType a = L"Text";
 *     xllType b = 20;
 *     xllType c = 30;
 *
 *     // Create array and return
 *     result = {a, b, c};
 *     result.push_back(L"Append item");
 *
 *     return result.get_return();
 * }
 * ```
 */
#define UDF(func, desc, ...)                                                                                 \
    Function func(__VA_ARGS__);                                                                              \
    struct func##_udf_register {                                                                             \
        func##_udf_register() {                                                                              \
            UDFRegistry::instance().registerFunction(__T(#func), count_args(&func))->set_info(EXPAND(desc)); \
        }                                                                                                    \
    } func##_udf_instance;                                                                                   \
    Function func(__VA_ARGS__)

 /// @brief Configure UDF function @param func Function name
#define UDFCONFIG(func) UDFRegistry::instance(__T(#func)).get_this()

/// @brief UDF configuration initialization function [Optional]
/// @note  Optional UDF settings, can also be configured by modifying source code
#define SET()                                   \
    int __udf_init();                              \
    struct __udf_init_register {                   \
        __udf_init_register() {                    \
            xll::init::instance().set(__udf_init); \
        }                                          \
    } __udf_init_instance;                         \
    int __udf_init()

#define EXPANDRTD(...) __VA_ARGS__
#define EXPAND_TO_PRIMITIVE(macro, args) macro args

/**
 * @brief Define and register RTD (Real-Time Data) function for creating Excel real-time data functions
 * @param func Function name, will be used as the callable function name in Excel
 * @param desc Function description, can be a string or map-like format configuration
 *             Supported formats: L"Simple description" or ({key, value}, {key, value}, ...)
 * @param rtdconfig RTD server configuration information, passed to RTDRegister for registration
 *                  Usually lambda expression or function pointer, defining actual data acquisition logic
 * @param ... Function parameter list, defines Excel parameters accepted by the RTD function
 *
 * * __Working Principle__:
 * 1. Declare and define an external export function (callable by Excel)
 * 2. Create static registration class instance, complete dual registration in constructor:
 *    - Register basic function information to UDFRegistry (name, parameters, help, etc.)
 *    - Register RTD specific configuration to RTDRegister (asynchronous processing, default values, etc.)
 * 3. Re-declare function, allowing implementation of specific function body after macro
 *
 * __RTD Configuration Parameters__:
 * `rtdconfig` should be a configuration containing the following elements:
 * - Lambda expression or function pointer: actual data acquisition logic
 * - Default value string: placeholder text displayed during function execution
 * - Asynchronous flag: true for asynchronous execution, false for synchronous execution
 *
 * __Calling in Excel__:
 * Use `=FunctionName(param1,param2,...)` in Excel cells to call
 *
 * @note RTD functions are especially suitable for data scenarios requiring frequent updates, such as:
 *       - Stock prices, exchange rates and other financial data
 *       - System monitoring data (CPU, memory usage)
 *       - File content monitoring
 *       - Network API data acquisition
 *
 * @warning Ensure RTD configuration lambda expressions or functions are thread-safe,
 *          as they may execute in RTD server worker threads
 *
 * @see RTDRegister RTD function registration manager
 * @see UDFRegistry UDF function registration manager
 * @see Topic RTD topic management class
 * @see CALLRTD RTD function call macro
 *
 * __Usage Example 1: Simple clock function__
 *
 * ```cpp
 * RTD(CurrentTime, L"Get current system time",
 *     ([](xllptrlist args, Topic* topic) {
 *         auto now = std::chrono::system_clock::now();
 *         auto time_t = std::chrono::system_clock::to_time_t(now);
 *         std::wstringstream ss;
 *         ss << std::put_time(std::localtime(&time_t), L"%Y-%m-%d %H:%M:%S");
 *         topic->setValue(ss.str());
 *         return 0;
 *     }, L"Loading...", true)) {
 *     xllType ret;
 *     CALLRTD(ret);
 *     return ret.get_return();
 * }
 * ```
 *
 * __Usage Example 2: File monitoring function with parameters__
 *
 * ```cpp
 * RTD(FileContent,
 *     ({udf::help, L"Monitor file content changes"},
 *      {udf::arguments, L"File path"}),
 *     ([](xllptrlist args, Topic* topic) {
 *         if (args.size() > 0) {
 *             std::wstring filepath = args[0]->get_str();
 *             std::wifstream file(filepath);
 *             if (file.is_open()) {
 *                 std::wstring content((std::istreambuf_iterator<wchar_t>(file)),
 *                                     std::istreambuf_iterator<wchar_t>());
 *                 topic->setValue(content);
 *             } else {
 *                 topic->setValue(L"File read failed");
 *             }
 *         }
 *         return 0;
 *     }, L"Reading...", true), Param filepath) {
 *     xllType ret;
 *     CALLRTD(ret, filepath);
 *     return ret.get_return();
 * }
 * ```
 *
 * __Usage Example 3: Stock price monitoring__
 *
 * ```cpp
 * RTD(StockPrice,
 *     ({udf::help, L"Get real-time stock price"},
 *      {udf::category, L"Financial Data"},
 *      {udf::arguments, L"Stock symbol"}),
 *     ([](xllptrlist args, Topic* topic) {
 *         if (args.size() > 0) {
 *             std::wstring symbol = args[0]->get_str();
 *             // This should call actual stock API
 *             double price = getStockPriceFromAPI(symbol);
 *             topic->setValue(std::to_wstring(price));
 *         }
 *         return 0;
 *     }, L"Getting price...", true), Param symbol) {
 *     xllType ret;
 *     CALLRTD(ret, symbol);
 *     return ret.get_return();
 * }
 * ```
 */
#define RTD(func, desc, rtdconfig, ...)                                                                        \
    Function func(__VA_ARGS__);                                                                                \
    struct func##_rtd_register {                                                                               \
        func##_rtd_register() {                                                                                \
            UDFRegistry::instance().registerFunction(__T(#func), count_args(&func))->set_info(EXPAND(desc));   \
            RTDRegister::instance().registerRTDFunction(__T(#func), EXPAND_TO_PRIMITIVE(EXPANDRTD, rtdconfig)); \
        }                                                                                                      \
    } func##_rtd_instance;                                                                                     \
    Function func(__VA_ARGS__)

 /// @brief Call RTD function, this method facilitates calling RTD functions
 /// @param ret Get return value
 /// @param ... Parameter list (try not to change the passed parameters as this may cause RTD service to fail to get results)
 /// @note This macro is prone to compilation failure. Under no parameter conditions, due to comma issues, some compilers will report errors when using this macro
 /// @note If compilation failure occurs, it is recommended to directly call the `xllRTD` function
#define CALLRTD(ret, ...) xllRTD(ret, __func__, ##__VA_ARGS__)