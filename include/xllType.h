/**
 * @file xllType.h
 * @brief Excel XLL 插件核心数据类型封装
 * @author mwm
 * @version 1.0
 * @date 2024-05-10
 * @copyright Copyright (c) 2024
 * 
 * 本文件定义了 ExcelXLL 框架的核心数据类型 xllType 类，该类继承自 Excel C API 的 xloper12 结构，
 * 为 Excel XLL 插件开发提供了类型安全、内存自动管理和便捷操作的高级接口。
 * 
 * __主要特性__：
 * - __类型安全__：提供强类型接口，避免直接操作底层 xloper12 结构
 * - __内存自动管理__：基于 RAII 原则，自动处理字符串和数组内存的分配与释放
 * - __丰富的类型转换__：支持数字、字符串、数组等多种 Excel 数据类型之间的转换
 * - __数组操作支持__：提供迭代器和索引访问，支持一维和二维数组操作
 * - __单元格引用处理__：自动处理 Excel 单元格引用和循环引用检测
 * - __序列化支持__：支持复杂数据结构的序列化和反序列化
 * 
 * __设计模式__：
 * - __继承模式__：继承 xloper12，保持与 Excel C API 的完全兼容性
 * - __RAII 模式__：构造函数分配资源，析构函数自动释放资源
 * - __迭代器模式__：提供标准的 begin/end 迭代器接口
 * - __智能指针管理__：使用 unique_ptr 管理动态分配的内存
 * 
 * __使用场景__：
 * - UDF 函数参数处理和返回值构造
 * - Excel 数据类型之间的安全转换
 * - 复杂数组和矩阵的操作处理
 * - RTD 函数的数据传递和处理
 * 
 * @warning 该类对象在 UDF 函数返回后会被 Excel 自动释放，请确保正确使用 get_return() 方法
 * @note 支持移动语义，可以高效地传递大型数组数据
 * 
 * @see xloper12 Excel C API 基础数据结构
 * @see UDF Excel 用户定义函数宏
 * @see RTD Excel 实时数据函数
 */
#pragma once
#include "XLCALL.H"
#include <vector>
#include <string>
#include <memory>

class xllType;

/// @brief xloper12 智能指针类型，用于自动内存管理
using xlptr = std::unique_ptr<xloper12>;

/// @brief xllType 智能指针类型，用于管理 xllType 对象
using xllptr = std::unique_ptr<xllType>;

/// @brief xllType 智能指针列表，用于管理 xllType 对象数组
using xllptrlist = std::vector<xllptr>;

/// @brief xllType 对象列表，用于构造一维数组
using xlllist = std::vector<xllType>;

/// @brief xllType 二维矩阵，用于构造二维数组
using xllmartix = std::vector<xlllist>;

/**
 * @brief Excel XLL 插件核心数据类型封装类
 * 
 * xllType 类是 ExcelXLL 框架的核心组件，继承自 Excel C API 的 xloper12 结构，
 * 为 Excel XLL 插件开发提供了高级、类型安全的数据封装接口。
 * 
 * __核心特性__：
 * 
 * __类型安全__：
 * - 提供强类型接口，避免直接操作复杂的 xloper12 结构
 * - 自动类型检测和转换，支持 is_num(), is_str(), is_array() 等方法
 * - 防止类型错误和数据损坏
 * 
 * __内存管理__：
 * - 基于 RAII 原则，构造时自动分配，析构时自动释放
 * - 使用 std::unique_ptr 管理动态内存，防止内存泄漏
 * - 支持移动语义，高效处理大数据传递
 * 
 * __数据类型支持__：
 * - __数值类型__：支持 double 类型的数值存储和转换
 * - __字符串类型__：支持 Unicode 字符串，兼容 std::wstring 和 C 风格字符串
 * - __数组类型__：支持一维和二维数组，提供迭代器和索引访问
 * - __单元格引用__：支持 Excel 单元格引用类型，自动解析引用内容
 * - __错误类型__：支持 Excel 错误类型，提供错误信息传递
 * 
 * __高级功能__：
 * - __循环引用检测__：防止 UDF 函数中的循环引用问题
 * - __数据序列化__：支持复杂数据结构的序列化和反序列化
 * - __迭代器支持__：提供 STL 风格的迭代器，支持 range-based for 循环
 * - __多维数组访问__：支持 at(row, col) 方式访问二维数组元素
 * 
 * __使用示例__：
 * 
 * ```cpp
 * // 1. 基本数值操作
 * xllType num(3.14);
 * xllType result = num.get_num() * 2;
 * 
 * // 2. 字符串处理
 * xllType str(L"你好世界");
 * std::wstring content = str.get_str();
 * 
 * // 3. 数组操作
 * xllType arr;
 * arr = {xllType(1.0), xllType(2.0), xllType(3.0)};
 * for (auto item : arr) {
 *     double val = item->get_num();
 * }
 * 
 * // 4. UDF 函数中使用
 * UDF(MyFunction, L"测试函数", Param input) {
 *     xllType result;
 *     xllType input_data = input;
 *     if (input_data.is_num()) {
 *         result = input_data.get_num() * 2;
 *     } else {
 *         result = L"类型错误";
 *     }
 *     return result.get_return();
 * }
 * ```
 * 
 * @note 继承自 xloper12，保持与 Excel C API 的完全兼容性
 * @warning 在 UDF 函数中使用时，必须通过 get_return() 方法返回结果
 * @warning 避免在多线程环境中同时修改同一个 xllType 对象
 * 
 * @see xloper12 Excel C API 基础数据结构
 * @see UDF 用户定义函数宏
 * @see xlltools.h 工具函数库
 */
class xllType : public xloper12 {
private:
    /// @brief 当前数组访问索引，用于记录最后一次访问的数组位置
    int pi = 0;
    
    /// @brief 数组行数，用于二维数组的行维度管理
    int rows = 0;
    
    /// @brief 数组列数，用于二维数组的列维度管理
    int cols = 0;
    
    /// @brief 错误代码，记录最后一次操作的错误信息
    int error_code = 0;
    
    /// @brief 数值存储，用于缓存从 xloper12 中解析出的数值
    double num = 0;
    
    /// @brief Excel 数据智能指针，用于管理从 Excel 获取的临时数据
    xlptr optr = nullptr;
    
    /// @brief 字符串存储，用于缓存从 xloper12 中解析出的字符串数据
    std::wstring str;
    
    /// @brief 数组元素存储，使用智能指针管理数组元素的内存
    xllptrlist array;
    
    /**
     * @brief 初始化对象为空状态
     * @return xllType* 返回当前对象指针以支持链式调用
     * 
     * 将对象的 xltype 设置为 xltypeNil，清空所有数据成员。
     * 该方法在构造函数中被调用，确保对象初始状态的一致性。
     */
    xllType* init();
    
    /**
     * @brief 从 xloper12 结构初始化对象
     * @param x 源 xloper12 结构引用
     * @return xllType* 返回当前对象指针以支持链式调用
     * 
     * 复制 xloper12 结构的 xltype 和 val 成员，然后调用 load() 方法
     * 解析具体的数据内容。该方法用于处理从 Excel 传入的数据。
     */
    xllType* init(const xloper12& x);
    
    /**
     * @brief 销毁对象并释放所有资源
     * @return xllType* 返回当前对象指针以支持链式调用
     * 
     * 清空字符串、数组和智能指针成员，释放所有动态分配的内存。
     * 在析构函数和重新赋值操作中被调用，确保内存安全。
     */
    xllType* destory();
    
    /**
     * @brief 加载和解析 xloper12 数据内容
     * @return xllType* 返回当前对象指针以支持链式调用
     * 
     * 根据 xloper12 的类型，解析并加载对应的数据内容。
     * 对于单元格引用类型，会调用 Excel API 获取实际数据。
     * 对于数组类型，会检查循环引用并解析数组元素。
     */
    xllType* load();
    
    /**
     * @brief 复制另一个 xllType 对象的所有数据
     * @param other 源 xllType 对象引用
     * @return xllType* 返回当前对象指针以支持链式调用
     * 
     * 执行深度复制，包括基本数据成员和数组元素。
     * 数组中的每个元素都会创建新的 xllType 对象副本。
     * 该方法在复制构造函数中被使用。
     */
    xllType* copy(const xllType& x);
    
    /**
     * @brief 加载单元格引用数据
     * @param type 需要转换的目标类型（xltypeNum, xltypeStr, xltypeMulti 等）
     * @return bool 成功返回 true，失败返回 false
     * 
     * 使用 Excel12(xlCoerce) API 将单元格引用转换为指定类型的数据。
     * 对于数组类型，会解析所有数组元素并存储到 array 成员中。
     * 失败时会记录错误代码到 error_code 成员。
     */
    bool load_ref(DWORD type);
    
    /**
     * @brief 检查循环引用
     * @return bool 未检测到循环引用返回 true，否则返回 false
     * 
     * 通过获取当前调用单元格的位置信息，检查是否与引用区域重叠。
     * 如果检测到循环引用，可能会导致 Excel 计算错误或死循环。
     * 该方法用于防止 UDF 函数中的无限递归问题。
     */
    bool check_ref();
public:
    
    /// @name 构造和析构函数
    
    /**
     * @brief 默认构造函数
     * 
     * 创建一个空的 xllType 对象，初始化为 xltypeNil 类型。
     * 所有数据成员都被设置为默认值，可以通过赋值操作符进行后续初始化。
     * 
     * __使用示例__：
     * ```cpp
     * xllType obj;           // 创建空对象
     * obj = 3.14;           // 赋值为数字
     * obj = L"你好";        // 赋值为字符串
     * ```
     */
    xllType();
    
    /**
     * @brief 析构函数
     * 
     * 自动释放所有动态分配的资源，包括字符串内存和数组元素。
     * 基于 RAII 原则，确保在对象生命周期结束时自动清理资源，
     * 防止内存泄漏和资源泄漏问题。
     * 
     * @note 析构函数会自动调用，不需要手动调用
     */
    ~xllType();
    
    /**
     * @brief 复制构造函数
     * @param other 要复制的源对象
     * 
     * 执行深度复制，创建一个与源对象完全独立的副本。
     * 对于数组类型，会递归复制所有元素，确保数据独立性。
     * 
     * __使用示例__：
     * ```cpp
     * xllType original(3.14);
     * xllType copy(original);   // 复制构造
     * copy = 2.71;             // 修改副本不影响原对象
     * ```
     */
    xllType(const xllType& other);
    
    /**
     * @brief 移动构造函数
     * @param other 要移动的源对象（右值引用）
     * 
     * 高效地转移源对象的资源所有权，避免不必要的复制操作。
     * 适用于大型数组或复杂数据结构的高效传递。
     * 移动后源对象将处于安全的空状态。
     * 
     * __使用示例__：
     * ```cpp
     * xllType createArray() {
     *     xllType arr;
     *     arr = {xllType(1), xllType(2), xllType(3)};
     *     return arr;  // 自动触发移动构造
     * }
     * xllType result = createArray();
     * ```
     */
    xllType(xllType&& other) noexcept;
    
    /**
     * @brief 从 xloper12 指针构造
     * @param px 指向 xloper12 结构的指针
     * 
     * 从 Excel 传入的 LPXLOPER12 指针创建 xllType 对象。
     * 主要用于 UDF 函数参数的初始化，自动解析 Excel 传入的数据。
     * 支持所有 Excel 数据类型，包括数值、字符串、数组和单元格引用。
     * 
     * __使用示例__：
     * ```cpp
     * UDF(MyFunction, L"测试", Param input) {
     *     xllType data(input);  // 从 Excel 参数创建
     *     return data.get_return();
     * }
     * ```
     * 
     * @warning 确保传入的指针非空且有效
     */
    xllType(const xloper12* px);
    
    /**
     * @brief 从 xloper12 结构构造
     * @param x xloper12 结构引用
     * 
     * 从 Excel 的 xloper12 结构创建 xllType 对象。
     * 与指针版本类似，但可以接受结构引用参数。
     * 自动识别和解析不同类型的 Excel 数据。
     * 
     * __使用示例__：
     * ```cpp
     * xloper12 excelData;
     * // ... 初始化 excelData
     * xllType data(excelData);
     * ```
     */
    xllType(const xloper12& x);
    
    /**
     * @brief 从数值构造
     * @param num 双精度浮点数
     * 
     * 创建一个包含指定数值的 xllType 对象。
     * 对象类型会被设置为 xltypeNum，支持所有 C++ 数值类型的隐式转换。
     * 
     * __使用示例__：
     * ```cpp
     * xllType pi(3.14159);
     * xllType count(42);
     * xllType result = pi.get_num() * count.get_num();
     * ```
     */
    xllType(double num);
    
    /**
     * @brief 从 C 风格字符串构造
     * @param str 以 null 结尾的 ANSI 字符串
     * 
     * 创建一个包含指定字符串的 xllType 对象。
     * ANSI 字符串会被自动转换为 Unicode 字符串存储。
     * 对象类型会被设置为 xltypeStr。
     * 
     * __使用示例__：
     * ```cpp
     * xllType text("Hello World");
     * std::wstring content = text.get_str();
     * ```
     * 
     * @note 推荐使用 Unicode 版本以获得更好的国际化支持
     */
    xllType(const char* str);
    
    /**
     * @brief 从 Unicode 字符串构造
     * @param ws 以 null 结尾的 Unicode 字符串
     * 
     * 创建一个包含指定 Unicode 字符串的 xllType 对象。
     * 支持所有 Unicode 字符，包括中文、特殊符号等。
     * 对象类型会被设置为 xltypeStr。
     * 
     * __使用示例__：
     * ```cpp
     * xllType chinese(L"你好世界");
     * xllType symbol(L"παβγ");
     * ```
     */
    xllType(const wchar_t* ws);
    
    /**
     * @brief 从 std::wstring 构造
     * @param str std::wstring 字符串引用
     * 
     * 创建一个包含指定 std::wstring 字符串的 xllType 对象。
     * 提供与 STL 字符串类的无缝集成，支持所有 std::wstring 的特性。
     * 对象类型会被设置为 xltypeStr。
     * 
     * __使用示例__：
     * ```cpp
     * std::wstring msg = L"动态生成的消息";
     * xllType result(msg);
     * ```
     */
    xllType(const std::wstring& str);

    xllType(const xllmartix &m);
    xllType(const xlllist &l);
    
    
    /// @name 赋值运算符
    
    /**
     * @brief 复制赋值运算符
     * @param other 要复制的源对象
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 执行深度复制，先清空当前对象的所有数据，然后复制源对象的所有内容。
     * 支持自赋值检测，对于数组类型会递归复制所有元素。
     * 
     * __使用示例__：
     * ```cpp
     * xllType a(3.14), b(2.71);
     * a = b;                    // 复制赋值
     * xllType c = a = b;        // 链式赋值
     * ```
     */
    xllType& operator=(const xllType& other);
    
    /**
     * @brief 从 xloper12 赋值
     * @param x 源 xloper12 结构引用
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 清空当前对象并从 xloper12 结构重新初始化。
     * 自动识别和解析不同类型的 Excel 数据，支持单元格引用的自动解析。
     * 
     * __使用示例__：
     * ```cpp
     * xloper12 excelData;
     * // ... 初始化 excelData
     * xllType obj;
     * obj = excelData;          // 从 Excel 数据赋值
     * ```
     */
    xllType& operator=(const xloper12& x);
    
    /**
     * @brief 从数值赋值
     * @param num 双精度浮点数
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 将对象设置为数值类型，并存储指定的数值。
     * 清空之前的所有数据，重新初始化为数值类型。
     * 支持所有 C++ 数值类型的隐式转换。
     * 
     * __使用示例__：
     * ```cpp
     * xllType obj;
     * obj = 3.14;              // 赋值为浮点数
     * obj = 42;                // 赋值为整数（自动转换）
     * ```
     */
    xllType& operator=(double num);
    
    /**
     * @brief 从 Unicode 字符串赋值
     * @param ws 以 null 结尾的 Unicode 字符串
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 将对象设置为字符串类型，并存储指定的 Unicode 字符串。
     * 清空之前的所有数据，重新初始化为字符串类型。
     * 支持所有 Unicode 字符，包括中文、特殊符号等。
     * 
     * __使用示例__：
     * ```cpp
     * xllType obj;
     * obj = L"你好世界";        // Unicode 中文
     * obj = L"παβγ";           // 希腊字母
     * ```
     */
    xllType& operator=(const wchar_t* ws);
    
    /**
     * @brief 从 std::wstring 赋值
     * @param str std::wstring 字符串引用
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 将对象设置为字符串类型，并存储指定的 std::wstring 字符串。
     * 提供与 STL 字符串类的无缝集成，支持所有 std::wstring 特性。
     * 清空之前的所有数据，重新初始化为字符串类型。
     * 
     * __使用示例__：
     * ```cpp
     * std::wstring msg = L"动态生成";
     * xllType obj;
     * obj = msg;               // 从 STL 字符串赋值
     * ```
     */
    xllType& operator=(const std::wstring& str);
    
    /**
     * @brief 从一维数组赋值
     * @param list xllType 对象的一维列表
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 将对象设置为数组类型，并存储指定的一维数组数据。
     * 数组会被设置为 1 行 N 列的矩阵形式，每个元素都会被复制。
     * 清空之前的所有数据，重新初始化为数组类型。
     * 
     * __使用示例__：
     * ```cpp
     * xlllist data = {xllType(1.0), xllType(2.0), xllType(L"文本")};
     * xllType obj;
     * obj = data;              // 赋值为一维数组
     * ```
     */
    xllType& operator=(const xlllist& list);
    
    /**
     * @brief 从二维矩阵赋值
     * @param matrix xllType 对象的二维矩阵
     * @return xllType& 返回当前对象的引用以支持链式赋值
     * 
     * 将对象设置为数组类型，并存储指定的二维矩阵数据。
     * 矩阵的行数和列数会被正确记录，支持 at(row, col) 方式访问。
     * 清空之前的所有数据，重新初始化为数组类型。
     * 
     * __使用示例__：
     * ```cpp
     * xllmartix matrix = {
     *     {xllType(1), xllType(2)},
     *     {xllType(3), xllType(4)}
     * };
     * xllType obj;
     * obj = matrix;            // 赋值为二维矩阵
     * ```
     */
    xllType& operator=(const xllmartix& matrix);
    
    /// @name 设值函数
    
    /// @brief 设置错误代码
    /// @param err Excel错误代码(如xlerrRef, xlerrValue等)
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_err(int err);
    
    /// @brief 设置数值
    /// @param num 双精度浮点数
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_num(double num);
    
    /// @brief 设置Unicode字符串
    /// @param ws 以null结尾的Unicode字符串
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_str(const wchar_t* ws);
    
    /// @brief 设置std::wstring字符串
    /// @param ws std::wstring字符串引用
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_str(const std::wstring& ws);
    
    /// @brief 设置一维数组
    /// @param l xllType对象的一维列表
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_list(const xlllist& l);
    
    /// @brief 设置二维矩阵
    /// @param m xllType对象的二维矩阵
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* set_matrix(const xllmartix& m);
    
    
    /// @name 取值函数
    
    /// @brief 获取数值
    /// @return double 返回存储的数值
    double get_num() const;
    
    /// @brief 获取字符串
    /// @return std::wstring 返回字符串内容
    std::wstring get_str() const;
    
    /// @brief 获取C风格字符串指针
    /// @return const wchar_t* 返回指向字符串的指针
    const wchar_t* get_c_str() const;
    
    /// @brief 获取最后一次错误代码
    /// @return int 返回错误代码，正常为0
    int get_last_err() const;
    
    
    /// @name 类型判断函数
    
    /// @brief 判断是否为数值类型
    /// @return bool 是数值类型返回true，否则返回false
    bool is_num() const;
    
    /// @brief 判断是否为字符串类型
    /// @return bool 是字符串类型返回true，否则返回false
    bool is_str() const;
    
    /// @brief 判断是否为数组类型
    /// @return bool 是数组类型返回true，否则返回false
    bool is_array() const;
    
    /// @brief 判断是否为单元格引用类型
    /// @return bool 是单元格引用返回true，否则返回false
    bool is_sref() const;
    
    
    /// @name 序列化函数
    
    /// @brief 将数组序列化为字符串
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* serialize();
    
    /// @brief 从字符串反序列化为数组
    /// @return xllType* 返回当前对象指针以支持链式调用
    xllType* deserialize();
    
    /// @brief 将其转换成适合赋值成适合传递xloper12类型的变量
    /// @return xllType* 返回当前对象指针以支持链式调用
    /// @warning 该操作可能会导致字符串使用异常，通常在最后使用, 使用完后不建议对该对象进行常规操作
    xllType* to_xloper12();
    
    /// @brief 获取Excel返回值指针
    /// @return xloper12* 返回Excel可识别的xloper12指针
    /// @warning 返回的指针由Excel负责释放，不要手动delete
    xloper12* get_return();
    
    /**
     * @brief STL风格迭代器类
     * 
     * 提供符合STL标准的迭代器接口，支持range-based for循环。
     * 用于遍历xllType数组中的每个元素，返回xllType指针。
     */
    class Iter {
    private:
        /// @brief 当前迭代位置索引
        int _i;
        /// @brief 指向父xllType对象的指针
        xllType* _p;
    public:
        /// @brief 构造函数
        /// @param p 指向xllType对象的指针
        /// @param i 初始迭代位置
        Iter(xllType* p, int i) : _i(i), _p(p) {}
        
        /// @brief 不等于比较运算符
        bool operator!=(const Iter& r) const;
        
        /// @brief 解引用运算符
        /// @return xllType* 返回当前位置的xllType指针
        xllType* operator*() const;
        
        /// @brief 前置递增运算符
        const Iter& operator++();
    };
    
    
    /// @name 迭代器和数组访问函数
    
    /// @brief 获取迭代器起始位置
    Iter begin();
    
    /// @brief 获取迭代器结束位置
    Iter end();
    
    /// @brief 获取数组元素数量
    int size() const;
    
    /// @brief 通过索引访问数组元素
    /// @param i 索引位置
    /// @return xllType* 返回指定位置的元素指针
    xllType* at(int i);
    
    /// @brief 通过行列索引访问二维数组元素
    /// @param row 行索引（1基索引）
    /// @param col 列索引（1基索引）
    /// @return xllType* 返回指定位置的元素指针
    xllType* at(int row, int col);
    
    /// @brief 数组下标访问运算符
    /// @param i 索引位置
    /// @return xllType* 返回指定位置的元素指针
    xllType* operator[](int i);
    
    /// @brief 添加元素
    /// @param x 要添加的元素 
    /// @return xllType* 返回当前对象指针
    xllType* push_back(const xllType& x);
    xllType* push_back(double num);
    xllType* push_back(const wchar_t* ws);
    xllType* push_back(const std::wstring& ws);
    
};