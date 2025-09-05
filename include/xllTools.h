/**
 * @file xlltools.h
 * @brief Some utility functions for XLL
 * @author mwmi
 * @date 2024-05-13
 * @copyright Copyright (c) 2024-2025 mwmi
 */

#pragma once

#include "XLCALL.H"
#include <string>
#include <vector>

 /// @brief Get function parameter count @tparam ReturnType Function return type @tparam ...Args Function parameter types @param func Function pointer @return Function parameter count
template <typename ReturnType, typename... Args>
constexpr int count_args(ReturnType(*func)(Args...)) {
    constexpr int param_count = sizeof...(Args);
    return param_count;
}

/// @brief Create Excel string (created string must be delete[] released, otherwise memory leak) @param ws String @return String pointer
wchar_t* makeStr12(const wchar_t* ws);

/// @brief Create Excel string (created xll string must be delete[] released, otherwise memory leak) @param s Source string @return Created xll string
wchar_t* makeStr12(const char* s);

/// @brief Create Excel string (created string must be delete[] released, otherwise memory leak) @param ws String @return String pointer
wchar_t* makeStr12(const std::wstring& ws);

/// @brief Create Excel string (created xll string must be delete[] released, otherwise memory leak) @param x xll string @return Created xll string
wchar_t* makeStr12(const xloper12& x);

/// @brief Copy string @param s Source string @return Copied string
wchar_t* copyStr(const char* s);

/// @brief Copy string @param ws Source string @return Target string
wchar_t* copyStr12(const wchar_t* ws);

/// @brief Convert Excel string to normal string @param ws String @return Normal string pointer
wchar_t* unmakeStr12(const wchar_t* ws);

/// @brief Convert xll string type to normal string @param ws Variable to store normal string @param valstr xll string type
void unmakeStr12(std::wstring& ws, const wchar_t* valstr);

/// @brief Convert xll string type to normal string @param xloper xll string type @return Normal string
std::wstring unmakeStr12(xloper12* xloper);

/// @brief Create xll string type @param ws String (string must be in xll string format) @return xll string type
xloper12 makeXllStr(wchar_t* ws);

/// @brief Create xll integer type @param i Integer @return xll integer type
xloper12 makeXllInt(int i);

/// @brief Create xll error type @param err Error code @return xll error type
xloper12 makeXllError(int err);

/// @brief Create xll number type @param d Number @return xll number type
xloper12 makeXllNum(double d);

/// @brief Serialize 2D string array @param data 2D string array @return Serialized string
bool xllSerialize(const std::vector<std::vector<std::wstring>>& data, std::wstring& result);

/// @brief Deserialize string @param str Serialized string @param result Deserialized 2D string array @return Deserialization success or not
bool xllDeserialize(const std::wstring& str, std::vector<std::vector<std::wstring>>& result);