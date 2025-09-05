/**
 * @file xllmacros.h
 * @author mwm
 * @brief 简单定义一些宏函数
 * @version 0.2
 * @date 2025-08-30
 * @copyright Copyright (c) 2024
 */
#pragma once
 // 长文本
#define __T(x) L##x
/// @brief 函数头简写
#define Function extern "C" __declspec(dllexport) LPXLOPER12
/// @brief 参数名简写
#define Param LPXLOPER12

/// @brief 展开括号 @param X 函数参数
#define EXPAND(X) ESC(ISH X)
#define ISH(...) ISH { __VA_ARGS__ }
#define ESC(...) ESC_(__VA_ARGS__)
#define ESC_(...) VAN##__VA_ARGS__
#define VANISH

/**
 * @brief 定义并注册UDF(User Defined Function)用户自定义函数，用于创建Excel自定义函数
 * @param func 函数名，将作为Excel中可调用的函数名
 * @param desc 函数描述信息，可以是字符串或类似map的格式配置
 *             支持格式：`L"简单描述"` 或 `({key, value}, {key, value}, ...)`
 * @param ... 函数参数列表，定义UDF函数接受的Excel参数
 *
 * __工作原理__：
 * 1. 声明并定义一个外部导出函数（Excel可调用）
 * 2. 创建静态注册类实例，在构造函数中完成函数注册：
 *    - 向UDFRegistry注册函数基本信息（名称、参数数量、帮助等）
 *    - 通过set_info设置函数描述信息和配置
 * 3. 重新声明函数，允许在宏后实现具体的函数体
 *
 * __函数配置参数__：
 * `desc` 可以是简单字符串描述或复杂配置映射：
 * - 简单格式：`L"函数功能描述"`
 * - 复杂格式：
 * ```cpp
 * (
 * { udf::name , L"value1" }, { udf::arguments, L"value2" },
 * { udf::help , L"value3" }, { udf::args_help, L"value4" },
 * { udf::category , L"value5" }, {udf::registername, L"value6" },
 * { udf::type , L"value7" }
 * )
 * ```
 * - 支持的配置项：`help`(帮助)、`category`(分类)、`arguments`(参数说明)等
 *
 * __Excel中的调用__：
 * 在Excel单元格中使用 `=函数名(参数1,参数2,...)` 调用
 *
 * @note UDF函数特别适用于自定义计算逻辑，如：
 *       - 数学运算函数（加法、求和、统计）
 *       - 字符串处理函数（拼接、格式化）
 *       - 数据转换和验证函数
 *       - 业务逻辑计算函数
 *
 * @warning 确保函数参数类型正确处理，使用xllType类进行类型检查和转换
 *
 * @see UDFRegistry UDF函数注册管理器
 * @see UDFCONFIG UDF函数配置宏
 * @see xllType Excel数据类型封装类
 * @see SET UDF初始化配置宏
 *
 * __使用示例1：简单加法函数__
 *
 * ```cpp
 * UDF(Add, L"两数相加", Param a, Param b) {
 *     xllType result;
 *     xllType a_ = a;
 *     xllType b_ = b;
 *     if (a_.is_num() && b_.is_num()) {
 *         result = a_.get_num() + b_.get_num();
 *     } else {
 *         result = L"类型错误！";
 *     }
 *     return result.get_return();
 * }
 * ```
 *
 * __使用示例2：带复杂配置的字符串处理函数__
 *
 * ```cpp
 * UDF(Concat2,
 *     ({udf::help, L"拼接两个字符串"},
 *      {udf::category, L"文本函数"},
 *      {udf::arguments, L"文本1,文本2"}),
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
 * __使用示例3：数组处理函数__
 *
 * ```cpp
 * UDF(MySum, L"数组求和", Param a) {
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
 * __使用示例4：返回数组的函数__
 *
 * ```cpp
 * UDF(RetArray, L"返回测试数组") {
 *     xllType result;
 *     xllType a = L"文本";
 *     xllType b = 20;
 *     xllType c = 30;
 *
 *     // 创建数组并返回
 *     result = {a, b, c};
 *     result.push_back(L"追加项");
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

 /// @brief 配置UDF函数 @param func 函数名
#define UDFCONFIG(func) UDFRegistry::instance(__T(#func)).get_this()

/// @brief UDF配置初始化函数【可选】
/// @note  UDF可选设置，也可以通过修改源代码来进行配置
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
 * @brief 定义并注册RTD(Real-Time Data)函数，用于创建Excel实时数据函数
 * @param func 函数名，将作为Excel中可调用的函数名
 * @param desc 函数描述信息，可以是字符串或类似map的格式配置
 *             支持格式：L"简单描述" 或 ({key, value}, {key, value}, ...)
 * @param rtdconfig RTD服务器配置信息，传递给RTDRegister进行注册
 *                  通常是lambda表达式或函数指针，定义实际的数据获取逻辑
 * @param ... 函数参数列表，定义RTD函数接受的Excel参数
 *
 * * __工作原理__：
 * 1. 声明并定义一个外部导出函数（Excel可调用）
 * 2. 创建静态注册类实例，在构造函数中完成双重注册：
 *    - 向UDFRegistry注册函数基本信息（名称、参数、帮助等）
 *    - 向RTDRegister注册RTD特定配置（异步处理、默认值等）
 * 3. 重新声明函数，允许在宏后实现具体的函数体
 *
 * __RTD配置参数__：
 * `rtdconfig` 应该是一个包含以下元素的配置：
 * - Lambda表达式或函数指针：实际的数据获取逻辑
 * - 默认值字符串：函数执行期间显示的占位文本
 * - 异步标志：true表示异步执行，false表示同步执行
 *
 * __Excel中的调用__：
 * 在Excel单元格中使用 `=函数名(参数1,参数2,...)` 调用
 *
 * @note RTD函数特别适用于需要频繁更新的数据场景，如：
 *       - 股票价格、汇率等金融数据
 *       - 系统监控数据（CPU、内存使用率）
 *       - 文件内容监控
 *       - 网络API数据获取
 *
 * @warning 确保RTD配置的lambda表达式或函数是线程安全的，
 *          因为它们可能在RTD服务器的工作线程中执行
 *
 * @see RTDRegister RTD函数注册管理器
 * @see UDFRegistry UDF函数注册管理器
 * @see Topic RTD主题管理类
 * @see CALLRTD RTD函数调用宏
 *
 * __使用示例1：简单时钟函数__
 *
 * ```cpp
 * RTD(CurrentTime, L"获取当前系统时间",
 *     ([](xllptrlist args, Topic* topic) {
 *         auto now = std::chrono::system_clock::now();
 *         auto time_t = std::chrono::system_clock::to_time_t(now);
 *         std::wstringstream ss;
 *         ss << std::put_time(std::localtime(&time_t), L"%Y-%m-%d %H:%M:%S");
 *         topic->setValue(ss.str());
 *         return 0;
 *     }, L"加载中...", true)) {
 *     xllType ret;
 *     CALLRTD(ret);
 *     return ret.get_return();
 * }
 * ```
 *
 * __使用示例2：带参数的文件监控函数__
 *
 * ```cpp
 * RTD(FileContent,
 *     ({udf::help, L"监控文件内容变化"},
 *      {udf::arguments, L"文件路径"}),
 *     ([](xllptrlist args, Topic* topic) {
 *         if (args.size() > 0) {
 *             std::wstring filepath = args[0]->get_str();
 *             std::wifstream file(filepath);
 *             if (file.is_open()) {
 *                 std::wstring content((std::istreambuf_iterator<wchar_t>(file)),
 *                                     std::istreambuf_iterator<wchar_t>());
 *                 topic->setValue(content);
 *             } else {
 *                 topic->setValue(L"文件读取失败");
 *             }
 *         }
 *         return 0;
 *     }, L"读取中...", true), Param filepath) {
 *     xllType ret;
 *     CALLRTD(ret, filepath);
 *     return ret.get_return();
 * }
 * ```
 *
 * __使用示例3：股票价格监控__
 *
 * ```cpp
 * RTD(StockPrice,
 *     ({udf::help, L"获取实时股票价格"},
 *      {udf::category, L"金融数据"},
 *      {udf::arguments, L"股票代码"}),
 *     ([](xllptrlist args, Topic* topic) {
 *         if (args.size() > 0) {
 *             std::wstring symbol = args[0]->get_str();
 *             // 这里应该调用实际的股票API
 *             double price = getStockPriceFromAPI(symbol);
 *             topic->setValue(std::to_wstring(price));
 *         }
 *         return 0;
 *     }, L"获取价格中...", true), Param symbol) {
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

 /// @brief 调用RTD函数, 此方法便于调用RTD函数
 /// @param ret 获取返回值
 /// @param ... 参数列表(传入的参数尽可能的不要有改变，因为这会导致rtd服务无法获取到结果)
 /// @note 该宏容易编译失败,在无参数条件下,因为逗号的原因,使用该宏,有的编译器会报错
 /// @note 若遇到编译失败的问题，建议直接调用`xllRTD`函数
#define CALLRTD(ret, ...) xllRTD(ret, __func__, ##__VA_ARGS__)