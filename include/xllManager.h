/**
 * @file xllManager.h
 * @brief 简单封装一下Excel的xll接口
 * @author mwmi
 * @date 2024-05-13
 * @copyright Copyright (c) 2024 by mwmi, All rights reserved.
 */
#pragma once

#include "XLCALL.H"
#include "xllType.h"
#include "xllTools.h"
#include "xllUDF.h"
#include "xllRTD.h"
#include "xllMacros.h"

/// @brief 声明XLL函数指针类型 @return int
using XllFunc = int (*)();

namespace xll {

/// @brief xll入口函数注册类
struct init {
    static init& instance() { static init i; return i; }
    void set(auto& f) { _init = f; }
    void call() { _init != nullptr && (_init)(); }
private:
    XllFunc _init = nullptr;
};

/// @brief 定义Excel的启动运行函数 @return int
extern XllFunc open;

/// @brief 定义Excel的关闭运行函数 @return int
extern XllFunc close;

/// @brief 定义add函数 @return int
extern XllFunc add;

/// @brief 定义remove函数 @return int
extern XllFunc remove;

/// @brief 定义register函数 @return int
extern XllFunc autoregister;

/// @brief xll管理器中显示名字
extern std::wstring xllName;

/// @brief 设置xll的默认类别
extern std::wstring defaultCategory;

/// @brief 是否启用RTD服务(默认为true)
extern bool enableRTD;

/// @brief 显示提示框 @param msg 提示内容 @param title 标题 @return int
int MsgBox(const wchar_t* msg, const wchar_t* title = L"提示");

/// @brief 使用Excel内置提示框显示提示信息 @param msg 提示的内容
bool alert(const wchar_t* msg);

/// @brief 获取单元格信息 @param cellInfo 单元格信息 @return bool 是否成功获取
bool getCellInfomation(XLOPER12& cellInfo);

/// @brief 获取xll文件的全路径 @param path xll文件的全路径 @return bool 是否成功获取
bool getXLLFullPath(std::wstring& path);

/// @brief 获取Excel的句柄 @param hwnd Excel的句柄 @return bool 是否成功获取
bool getExcelHandle(HWND& hwnd);

/// @brief 根据Excel表达式返回值【注意：表达式不需要加等号】 @param expression Excel表达式 @param value 返回值 @return bool 是否成功获取
/// @note 表达式示例：SUM(A1:A5)
/// @note 注意返回结果不会随着单元格内容的改变而改变
bool evaluate(std::wstring expression, xllType& value);

/// @brief 调用Excel函数 @tparam ...Args 参数类型 @param result 返回值 @param xlfn 函数编号 @param ...args 参数 @return int 调用结果
/// @note 函数编号参考XLCALL.H中的Xlf
template <typename... Args>
int callExcelFunction(xllType& result, int xlfn, Args... args) {
    constexpr int n = sizeof...(args);
    int ret = xlretSuccess;
    xloper12 r;
    LPXLOPER12 v[n];
    xllType arr[] = {[&]() {
        return xllType(args);
    }()...};
    for (int i = 0; i < n; i++) {
        v[i] = arr[i].to_xloper12();
    }
    ret = Excel12v(xlfn, &r, n, v);
    if (ret == xlretSuccess) {
        result = r;
        Excel12(xlFree, 0, 1, &r);
    } else {
        result.set_err(xlerrValue);
    }
    return ret;
}
} // namespace xll