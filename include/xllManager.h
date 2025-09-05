/**
 * @file xllManager.h
 * @brief Simple encapsulation of Excel's xll interface
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

/// @brief Declare XLL function pointer type @return int
using XllFunc = int (*)();

namespace xll {

/// @brief xll entry function registration class
struct init {
    static init& instance() { static init i; return i; }
    void set(auto& f) { _init = f; }
    void call() { _init != nullptr && (_init)(); }
private:
    XllFunc _init = nullptr;
};

/// @brief Define Excel's startup function @return int
extern XllFunc open;

/// @brief Define Excel's close function @return int
extern XllFunc close;

/// @brief Define add function @return int
extern XllFunc add;

/// @brief Define remove function @return int
extern XllFunc remove;

/// @brief Define register function @return int
extern XllFunc autoregister;

/// @brief Display name in xll manager
extern std::wstring xllName;

/// @brief Set default category for xll
extern std::wstring defaultCategory;

/// @brief Whether to enable RTD service (default is true)
extern bool enableRTD;

/// @brief Show message box @param msg Message content @param title Title @return int
int MsgBox(const wchar_t* msg, const wchar_t* title = L"Tip");

/// @brief Use Excel built-in alert box to display message @param msg Message content
bool alert(const wchar_t* msg);

/// @brief Get cell information @param cellInfo Cell information @return bool Whether successfully obtained
bool getCellInfomation(XLOPER12& cellInfo);

/// @brief Get full path of xll file @param path Full path of xll file @return bool Whether successfully obtained
bool getXLLFullPath(std::wstring& path);

/// @brief Get Excel handle @param hwnd Excel handle @return bool Whether successfully obtained
bool getExcelHandle(HWND& hwnd);

/// @brief Return value based on Excel expression [Note: expression does not need equal sign] @param expression Excel expression @param value Return value @return bool Whether successfully obtained
/// @note Expression example: SUM(A1:A5)
/// @note Note that the return result will not change with cell content changes
bool evaluate(std::wstring expression, xllType& value);

/// @brief Call Excel function @tparam ...Args Parameter types @param result Return value @param xlfn Function number @param ...args Parameters @return int Call result
/// @note Function numbers refer to Xlf in XLCALL.H
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