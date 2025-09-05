/**
 * @file xlltools.h
 * @brief XLL 的一些工具函数
 * @author mwmi
 * @date 2024-05-13
 * @copyright Copyright (c) 2024-2025 mwmi
 */

#pragma once

#include <string>
#include <vector>

 /// @brief 取得函数参数个数 @tparam ReturnType 函数返回值类型 @tparam ...Args 函数参数类型 @param func 函数指针 @return 函数参数个数
template <typename ReturnType, typename... Args>
constexpr int count_args(ReturnType(*func)(Args...)) {
    constexpr int param_count = sizeof...(Args);
    return param_count;
}

/// @brief 创建Excel字符串(创建的字符串必须delete[]释放, 否则会导致内存泄漏) @param ws 字符串 @return 字符串指针
wchar_t* makeStr12(const wchar_t* ws);

/// @brief 创建Excel字符串(创建的xll字符串必须delete[]释放, 否则会导致内存泄漏) @param s 源字符串 @return 创建的xll字符串
wchar_t* makeStr12(const char* s);

/// @brief 创建Excel字符串(创建的字符串必须delete[]释放, 否则会导致内存泄漏) @param ws 字符串 @return 字符串指针
wchar_t* makeStr12(const std::wstring& ws);

/// @brief 创建Excel字符串(创建的xll字符串必须delete[]释放, 否则会导致内存泄漏) @param x xll字符串 @return 创建的xll字符串
wchar_t* makeStr12(const xloper12& x);

/// @brief 复制字符串 @param s 源字符串 @return 复制的字符串
wchar_t* copyStr(const char* s);

/// @brief 复制字符串 @param ws 源字符串 @return 目标字符串
wchar_t* copyStr12(const wchar_t* ws);

/// @brief 将Excel字符串转换成普通字符串 @param ws 字符串 @return 普通字符串指针
wchar_t* unmakeStr12(const wchar_t* ws);

/// @brief 将xll字符串类型转换成普通字符串 @param ws 存储普通字符串的变量 @param valstr xll字符串类型
void unmakeStr12(std::wstring& ws, const wchar_t* valstr);

/// @brief 将xll字符串类型转换成普通字符串 @param xloper xll字符串类型 @return 普通字符串
std::wstring unmakeStr12(xloper12* xloper);

/// @brief 创建xll字符串类型 @param ws 字符串(字符串必须为符合xll字符串格式的字符串) @return xll字符串类型
xloper12 makeXllStr(wchar_t* ws);

/// @brief 创建xll整数类型 @param i 整数 @return xll整数类型
xloper12 makeXllInt(int i);

/// @brief 创建xll错误类型 @param err 错误码 @return xll错误类型
xloper12 makeXllError(int err);

/// @brief 创建xll数字类型 @param d 数字 @return xll数字类型
xloper12 makeXllNum(double d);

/// @brief 序列化二维字符串数组 @param data 二维字符串数组 @return 序列化后的字符串
bool xllSerialize(const std::vector<std::vector<std::wstring>>& data, std::wstring& result);

/// @brief 反序列化字符串 @param str 序列化后的字符串 @param result 反序列化后的二维字符串数组 @return 反序列化成功与否
bool xllDeserialize(const std::wstring& str, std::vector<std::vector<std::wstring>>& result);