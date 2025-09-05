# 🚀 ExcelXLL - 现代 Excel 插件开发框架

<div align="center">

## 🤖 AI Generated Documentation 🤖
**✨ 本文档由人工智能生成 ✨**

---

</div>

[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue.svg)](https://en.cppreference.com/w/cpp/20)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/mwmi/excel-xll/blob/main/LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows/)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-orange.svg)](https://www.microsoft.com/excel)
[![WPS](https://img.shields.io/badge/WPS-Office-red.svg)](https://www.wps.com/)
[![GitHub Stars](https://img.shields.io/github/stars/mwmi/excel-xll?style=social)](https://github.com/mwmi/excel-xll)
[![GitHub Forks](https://img.shields.io/github/forks/mwmi/excel-xll?style=social)](https://github.com/mwmi/excel-xll/fork)

📈 ExcelXLL 是一个功能强大的 Excel/WPS XLL 插件开发框架，使用现代 C++20 标准构建。该项目旨在简化 Excel 自定义函数的开发过程，为 C++ 开发者提供高效、易用、内存安全的插件开发解决方案。

## ✨ 核心特性

### 💻 开发便利性
- **⚡ 现代 C++20 标准**：充分利用现代 C++ 特性提升开发效率
- **🎯 简洁的宏定义**：通过 `UDF` 和 `RTD` 宏大幅简化函数定义
- **🔄 智能类型系统**：自动处理 Excel 与 C++ 之间的数据类型转换
- **📚 丰富的示例代码**：提供完整的函数示例和最佳实践

### 🛡️ 内存安全
- **🔒 RAII 模式**：自动管理内存资源，防止内存泄漏
- **🧠 智能指针管理**：使用 `std::unique_ptr` 管理动态内存
- **⚠️ 异常安全保证**：提供强异常安全性保证

### 🚀 高性能
- **⚡ 零拷贝设计**：支持移动语义，高效处理大数据
- **📦 静态链接优化**：Release 模式下减少依赖和文件大小
- **🔧 编译时优化**：模板元编程和编译时计算

### 🌐 广泛兼容性
- **🔨 双编译器支持**：MinGW-w64 和 Microsoft Visual C++
- **💾 多架构支持**：32 位 WPS 和 64 位 Excel
- **📅 多版本兼容**：Excel 2016+ 和 WPS Office

### 🔄 实时数据支持
- **📡 RTD 函数**：支持实时数据更新
- **⚡ 异步处理**：非阻塞的数据获取和更新
- **🔗 COM 组件**：完整的 RTD 服务器实现

## 📁 项目结构图

```
ExcelXLL/
├── include/                 # 头文件目录
│   ├── XLCALL.H            # Excel C API 接口
│   ├── xllType.h           # 核心数据类型封装
│   ├── xllUDF.h            # UDF 函数管理
│   ├── xllManager.h        # XLL 管理器
│   ├── xllRTD.h            # RTD 函数管理
│   ├── xllTools.h          # 工具函数库
│   ├── xllMacros.h         # 宏定义
│   ├── RtdServer.h         # RTD 服务器
│   ├── RTDTopic.h          # RTD 主题管理
│   ├── IRTDServer.h        # RTD 服务器接口
│   └── dll.h               # DLL 导出定义
├── src/                    # 源文件目录
│   ├── XLCALL.CPP          # Excel API 实现
│   ├── xllType.cpp         # 数据类型实现
│   ├── xllUDF.cpp          # UDF 框架实现
│   ├── xllManager.cpp      # 管理器实现
│   ├── xllRTD.cpp          # RTD 实现
│   ├── xllTools.cpp        # 工具函数实现
│   ├── RtdServer.cpp       # RTD 服务器实现
│   ├── RTDTopic.cpp        # RTD 主题实现
│   └── dll.cpp             # DLL 入口实现
├── lib/                    # 库文件目录
│   ├── XLCALL32.LIB        # 32位 Excel 库
│   └── x64/
│       └── XLCALL32.LIB    # 64位 Excel 库
├── functions.cpp           # 示例函数实现
├── dll.def                 # DLL 导出定义
├── CMakeLists.txt          # CMake 构建配置
└── LICENSE                 # MIT 许可证
```

## 🚀 快速开始

### 📋 系统要求

- **💿 操作系统**：Windows 10/11
- **🔨 编译器**：
  - 🟢 MinGW-w64 GCC 11+ (推荐)
  - 🔵 Microsoft Visual C++ 2019+
- **⚙️ 构建工具**：CMake 3.20+, Ninja
- **🎯 目标平台**：
  - 📊 Excel 2016+ (64位)
  - 📝 WPS Office (32位)

### ⚙️ 环境配置

#### 🟢 方案一：MinGW-w64 (推荐)

1. **安装 MSYS2**
```bash
# 更新包管理器
pacman -Syu

# 安装工具链
pacman -S mingw-w64-i686-toolchain mingw-w64-x86_64-toolchain
pacman -S mingw-w64-x86_64-cmake mingw-w64-x86_64-ninja
```

2. **配置环境变量**
```
C:\msys64\mingw64\bin
C:\msys64\mingw32\bin
C:\msys64\usr\bin
```

#### 🔵 方案二：Microsoft Visual C++

安装 Visual Studio 2019/2022，确保包含以下组件：
- C++ 核心功能
- MSVC v143 编译器工具集
- Windows 10/11 SDK
- CMake 工具

### 🔨 构建项目

#### MinGW-w64 构建

```bash
# 32位构建 (适用于 WPS)
cmake -B build32 -G Ninja -DCMAKE_CXX_STANDARD=20 -DCMAKE_SIZEOF_VOID_P=4
cmake --build build32

# 64位构建 (适用于 Excel)
cmake -B build64 -G Ninja -DCMAKE_CXX_STANDARD=20 -DCMAKE_SIZEOF_VOID_P=8
cmake --build build64
```

#### MSVC 构建

```cmd
# 初始化 MSVC 环境
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"

# 32位构建
cmake -B build32 -G Ninja -DCMAKE_C_COMPILER=cl -DCMAKE_CXX_COMPILER=cl -DCMAKE_CXX_STANDARD=20 -DCMAKE_SIZEOF_VOID_P=4
cmake --build build32

# 64位构建
cmake -B build64 -G Ninja -DCMAKE_C_COMPILER=cl -DCMAKE_CXX_COMPILER=cl -DCMAKE_CXX_STANDARD=20 -DCMAKE_SIZEOF_VOID_P=8
cmake --build build64
```

### 📥 安装和使用

1. **复制插件文件**：
   - Excel：将 `build64/EXCELXLL_x64.xll` 复制到目标位置
   - WPS：将 `build32/EXCELXLL_x86.xll` 复制到目标位置

2. **加载插件**：
   - 打开 Excel/WPS
   - 文件 → 选项 → 加载项 → 管理Excel加载项 → 浏览
   - 选择对应的 `.xll` 文件

3. **使用函数**：
   ```excel
   =HelloWorld()           # 显示"你好世界"
   =Add(10, 20)           # 计算两数之和
   =MySum(A1:A10)         # 对范围求和
   =RTDClock()            # 显示实时时钟
   ```

## 📖 开发指南手册

### 🎯 创建 UDF 函数

UDF (User Defined Function) 用于创建 Excel 自定义函数：

```cpp
#include "xllManager.h"

// 简单函数示例
UDF(HelloWorld, L"问候函数") {
    xllType result;
    result = L"你好世界";
    return result.get_return();
}

// 带参数的函数
UDF(Add, L"两数相加", Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    
    if (a_.is_num() && b_.is_num()) {
        result = a_.get_num() + b_.get_num();
    } else {
        result = L"类型错误！";
    }
    return result.get_return();
}

// 复杂配置函数
UDF(Concat2, 
    ({udf::help, L"拼接两个字符串"},
     {udf::category, L"文本函数"},
     {udf::arguments, L"文本1,文本2"}),
    Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    std::wstring str = a_.get_str() + b_.get_str();
    result = str;
    return result.get_return();
}

// 数组处理函数
UDF(MySum, L"数组求和", Param a) {
    xllType result;
    xllType a_ = a;
    double sum = 0;
    
    if (a_.is_array()) {
        // 使用迭代器遍历
        for (auto item : a_) {
            if (item->is_num()) {
                sum += item->get_num();
            }
        }
    }
    result = sum;
    return result.get_return();
}

// 返回数组函数
UDF(RetArray, L"返回测试数组") {
    xllType result;
    xllType a = L"文本";
    xllType b = 20;
    xllType c = 30;
    xllType d = 40;
    
    // 创建二维数组
    result = {{a, b}, {c, d}};
    return result.get_return();
}
```

### 📡 创建 RTD 函数

RTD (Real-Time Data) 用于创建实时数据函数：

```cpp
// 简单 RTD 函数
RTD(RTDHelloWorld, L"RTD问候函数", 
    ([](xllptrlist args, Topic* topic) {
        topic->setValue(L"Hello World");
        return 0;
    }, L"加载中...")) {
    
    xllType ret;
    CALLRTD(ret);
    return ret.get_return();
}

// 异步 RTD 时钟
RTD(RTDClock, L"实时时钟", 
    ([](xllptrlist args, Topic* topic) {
        SYSTEMTIME st;
        while (true) {
            GetLocalTime(&st);
            wchar_t buffer[100];
            swprintf(buffer, 100, L"🕒 %04d-%02d-%02d %02d:%02d:%02d", 
                    st.wYear, st.wMonth, st.wDay, 
                    st.wHour, st.wMinute, st.wSecond);
            topic->setValue(buffer);
            Sleep(1000);
        }
        return 0;
    }, L"时钟启动中...", true)) {
    
    xllType ret;
    CALLRTD(ret);
    return ret.get_return();
}

// 带参数的 RTD 函数
RTD(RTDParam, L"参数测试", 
    ([](xllptrlist args, Topic* topic) {
        xllType p;
        if (args.size() > 0) {
            p = *args[0].get();
        }
        topic->setValue(p);
        return 0;
    }, L"参数处理中..."), Param a) {
    
    xllType ret;
    CALLRTD(ret, a);
    return ret.get_return();
}
```

### ⚙️ 全局配置

```cpp
// 插件配置
SET() {
    // 设置插件名称
    xll::xllName = L"我的Excel插件";
    
    // 设置默认函数分类
    xll::defaultCategory = L"自定义功能";
    
    // 设置是否启用RTD (默认启用)
    xll::enableRTD = true;
    
    // 插件加载时回调
    xll::open = []() {
        // 配置特定函数
        UDFCONFIG(HelloWorld)->set_funchelp(L"显示问候信息");
        
        // 显示欢迎消息
        xll::alert(L"插件加载成功！");
        return 1;
    };
    
    // 插件卸载时回调
    xll::close = []() {
        return 1;
    };
    
    return 0;
}
```

### 🔄 数据类型处理

`xllType` 类提供了强大的数据类型封装：

```cpp
// 类型检查
if (param.is_num()) { /* 数值类型 */ }
if (param.is_str()) { /* 字符串类型 */ }
if (param.is_array()) { /* 数组类型 */ }

// 数据获取
double number = param.get_num();
std::wstring text = param.get_str();

// 数组操作
for (auto item : param) {
    if (item->is_num()) {
        double val = item->get_num();
    }
}

// 索引访问
auto item = param[0];      // 线性索引
auto item = param.at(0, 1); // 二维索引

// 数据设置
xllType result;
result = 3.14;           // 数值
result = L"文本";        // 字符串
result = {1, 2, 3};      // 一维数组
result = {{1, 2}, {3, 4}}; // 二维数组
```

## 🔧 高级功能特性

### 📞 调用 Excel 内置函数

```cpp
UDF(CallExcelFunc, L"调用Excel函数示例", Param range) {
    xllType result;
    xllType range_ = range;
    
    // 调用Excel的SUM函数
    xll::callExcelFunction(result, xlfSum, range_);
    
    // 调用Excel的ABS函数
    // xll::callExcelFunction(result, xlfAbs, -100);
    
    return result.get_return();
}
```

### ⚠️ 错误处理机制

```cpp
UDF(SafeDiv, L"安全除法", Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    
    if (!a_.is_num() || !b_.is_num()) {
        result.set_err(xlerrValue);  // #VALUE! 错误
        return result.get_return();
    }
    
    double divisor = b_.get_num();
    if (divisor == 0) {
        result.set_err(xlerrDiv0);   // #DIV/0! 错误
        return result.get_return();
    }
    
    result = a_.get_num() / divisor;
    return result.get_return();
}
```

### ⚡ 性能优化技巧

1. **避免不必要的复制**：
```cpp
// 好的做法
UDF(ProcessArray, L"处理数组", Param arr) {
    xllType result;
    xllType& arr_ref = arr;  // 引用，避免复制
    // ... 处理逻辑
    return result.get_return();
}
```

2. **使用移动语义**：
```cpp
xllType createLargeArray() {
    xllType arr;
    // ... 创建大数组
    return arr;  // 自动使用移动语义
}
```

3. **预分配内存**：
```cpp
UDF(CreateArray, L"创建数组", Param size) {
    xllType result;
    int n = static_cast<int>(xllType(size).get_num());
    
    result.reserve(n);  // 预分配内存
    for (int i = 0; i < n; ++i) {
        result.push_back(i);
    }
    
    return result.get_return();
}
```

## 🐛 故障排除指南

### ❓ 常见问题

1. **编译错误：找不到 XLCALL32.LIB**
   - 确保 `lib/` 目录包含正确的库文件
   - 检查 CMake 配置中的 `link_directories`

2. **Excel 无法加载 XLL 文件**
   - 检查架构匹配（32位 vs 64位）
   - 确保所有依赖库都已正确链接
   - 检查 Windows 安全设置

3. **函数返回 #VALUE! 错误**
   - 检查参数类型是否正确
   - 确保正确使用 `get_return()` 方法
   - 检查内存管理是否正确

4. **RTD 函数不更新**
   - 确保 RTD 服务器已正确注册
   - 检查异步函数是否正常运行
   - 验证 COM 组件注册

### 🔍 调试技巧

1. **使用断点调试**：
```cpp
UDF(DebugFunc, L"调试函数", Param input) {
    xllType result;
    
    // 在此处设置断点
    xllType input_ = input;
    
    if (input_.is_num()) {
        result = input_.get_num() * 2;
    }
    
    return result.get_return();
}
```

2. **输出调试信息**：
```cpp
UDF(LogFunc, L"日志函数", Param input) {
    xllType result;
    
    // 使用Excel提示框显示调试信息
    std::wstring debug_info = L"输入类型: " + 
        (xllType(input).is_num() ? L"数值" : L"非数值");
    xll::alert(debug_info.c_str());
    
    result = L"调试完成";
    return result.get_return();
}
```

### 📊 性能分析

1. **使用 Release 模式构建**以获得最佳性能
2. **避免频繁的字符串操作**
3. **合理使用 RTD 更新频率**
4. **考虑使用线程池**处理复杂计算

## 🤝 贡献指南说明

我们欢迎任何形式的贡献！请遵循以下步骤：

1. **Fork 项目**
2. **创建功能分支** (`git checkout -b feature/AmazingFeature`)
3. **提交更改** (`git commit -m 'Add some AmazingFeature'`)
4. **推送到分支** (`git push origin feature/AmazingFeature`)
5. **开启 Pull Request**

### 📝 代码规范

- 使用 C++20 标准
- 遵循现有的代码风格
- 添加适当的注释和文档
- 编写单元测试
- 确保所有测试通过

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 特别致谢

- Microsoft Excel SDK 团队
- MinGW-w64 项目
- CMake 项目
- 所有贡献者和用户

## 📞 联系方式

- 👨‍💻 **作者**：mwmi
- 📧 **邮箱**：[通过 GitHub Issues 联系](https://github.com/mwmi/excel-xll/issues)
- 🔗 **项目链接**：[https://github.com/mwmi/excel-xll](https://github.com/mwmi/excel-xll)
- 🌟 **GitHub**：[@mwmi](https://github.com/mwmi)

---

<div align="center">

### 🌟 如果这个项目对您有帮助，请给我们一个 Star！

[![GitHub stars](https://img.shields.io/github/stars/mwmi/excel-xll.svg?style=social&label=Star)](https://github.com/mwmi/excel-xll)
[![GitHub forks](https://img.shields.io/github/forks/mwmi/excel-xll.svg?style=social&label=Fork)](https://github.com/mwmi/excel-xll/fork)
[![GitHub watchers](https://img.shields.io/github/watchers/mwmi/excel-xll.svg?style=social&label=Watch)](https://github.com/mwmi/excel-xll)

📚 **更多文档和示例请查看** [📖 Wiki](https://github.com/mwmi/excel-xll/wiki)

🚀 **快速开始** | 📖 **详细文档** | 🐛 **问题反馈** | 💡 **功能建议**

[开始使用](https://github.com/mwmi/excel-xll#-快速开始) • [查看文档](https://github.com/mwmi/excel-xll#-开发指南手册) • [提交问题](https://github.com/mwmi/excel-xll/issues) • [功能请求](https://github.com/mwmi/excel-xll/issues/new)

**感谢所有贡献者的支持！** 🎉

</div>