/**
 * @file xlludf.h
 * @brief Define XLL UDF function management class
 * @author mwmi
 * @date 2024-05-14
 * @copyright Copyright (c) 2024, mwmi
 */
#pragma once

#include <map>
#include <string>

namespace udf {
/// @brief UDF function information structure
enum function {
  /// @brief Function name displayed in Excel
  name,
  /// @brief Help text of the function displayed in Excel
  help,
  /// @brief Formula category of the function displayed in Excel
  category,
  /// @brief Return type and parameter types of the function
  type,
  /// @brief Parameter text of the function displayed in Excel
  arguments,
  /// @brief Parameter description text of the function displayed in Excel
  args_help,
  /// @brief Name of the currently written function
  registername,
};
}

struct UDFInfo {
  /// @brief Number of parameters
  int paramNum;
  /// @brief Function name declared in file
  wchar_t* register_name = nullptr;
  /// @brief Return type and parameter types of the function
  wchar_t* type_text = nullptr;
  /// @brief Function name displayed in Excel
  wchar_t* function_name = nullptr;
  /// @brief Parameter text of the function displayed in Excel
  wchar_t* argument_text = nullptr;
  /// @brief Formula category of the function displayed in Excel
  wchar_t* category = nullptr;
  /// @brief Description text of the function displayed in Excel
  wchar_t* function_help = nullptr;
  /// @brief Parameter description text of the function displayed in Excel
  wchar_t* argument_help = nullptr;
};

/// @brief Function information structure
using UDFList = std::map<std::wstring, UDFInfo>;

/// @brief UDF registration class
/// @details The UDFRegistry class can register UDF functions in XLL
class UDFRegistry {
public:
  /// @brief Get current object @param name Name of the registration function (optional) @return Current object
  static UDFRegistry& instance(std::wstring name = L"");

  /// @brief Register UDF function @param name Name of the registration function @param paramNum Number of parameters of the registration function @return Current object
  UDFRegistry* registerFunction(const std::wstring& name, const int& paramNum);

  /// @brief Register function
  UDFRegistry* regist();

  /// @brief Unregister function
  UDFRegistry* unregist();

  /// @brief Automatically register all functions
  UDFRegistry* AutoRegist();

  /// @brief Automatically unregister all functions
  UDFRegistry* AutoUnRegist();

  /// @brief Get current object @return Current object
  UDFRegistry* get_this();

  /// @brief Set help content of function in Excel function @param info Function information @return Current object
  UDFRegistry* set_info(const wchar_t* info);

  /// @brief Batch set function information @param info Function information dictionary @return Current object
  /// @note Function information format: {
  /// @note {udf::name, "value1"},{udf::help, "value2"},
  /// @note {udf::category, "value3"},{udf::type, "value4"},
  /// @note {udf::arguments, "value5"},{udf::args_help, "value6"}
  /// @note {udf::registername, "value7"}
  /// @note }
  UDFRegistry* set_info(std::map<udf::function, std::wstring> info);

  /// @brief Set name of registration function @param name Registration name of the function @return Return current object
  UDFRegistry* set_regsitername(const wchar_t* name);

  /// @brief Set function name in Excel @param name Function name @return Return current object
  UDFRegistry* set_funcname(const wchar_t* name);

  /// @brief Set parameter types of function in Excel @param text Parameter type placeholder like "UUU" @return Return current object
  UDFRegistry* set_typetext(const wchar_t* text);

  /// @brief Set function parameter prompts in Excel @param text Parameter names (like: "param1,param2,param3" separated by commas) @return Return current object
  UDFRegistry* set_argstip(const wchar_t* text);

  /// @brief Set formula category of function in Excel functions @param category Category name @return Return current object
  UDFRegistry* set_category(const wchar_t* category);

  /// @brief Set help content of function in Excel functions @param help Help documentation @return Return current object
  UDFRegistry* set_funchelp(const wchar_t* help);

  /// @brief Set help content for each parameter of function in Excel @param help Help documentation (like: "param1,param2,param3" separated by commas) @return Return current object
  UDFRegistry* set_argshelp(const wchar_t* help);

private:
  std::wstring name;
  UDFList udfs;
};