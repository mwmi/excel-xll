/**
 * @file xlludf.h
 * @brief 定义XLL UDF函数管理类
 * @author mwmi
 * @date 2024-05-14
 * @copyright Copyright (c) 2024, mwmi
 */
#pragma once

#include <map>
#include <string>

namespace udf {
/// @brief UDF函数信息结构体
enum function {
  /// @brief Excel中显示的函数的名字
  name,
  /// @brief Excel中显示的函数的帮助文本
  help,
  /// @brief Excel中显示的函数的公式类别
  category,
  /// @brief 函数的返回类型和参数类型
  type,
  /// @brief Excel中显示的函数的参数文本
  arguments,
  /// @brief Excel中显示的函数的参数说明文本
  args_help,
  /// @brief 当前编写函数的名字
  registername,
};
}

struct UDFInfo {
  /// @brief 参数个数
  int paramNum;
  /// @brief 文件中声明的函数名字
  wchar_t* register_name = nullptr;
  /// @brief 函数的返回类型和参数类型
  wchar_t* type_text = nullptr;
  /// @brief Excel中显示的函数的名字
  wchar_t* function_name = nullptr;
  /// @brief Excel中显示的函数的参数文本
  wchar_t* argument_text = nullptr;
  /// @brief Excel中显示的函数的公式类别
  wchar_t* category = nullptr;
  /// @brief Excel中显示的函数的说明文本
  wchar_t* function_help = nullptr;
  /// @brief Excel中显示的函数的参数说明文本
  wchar_t* argument_help = nullptr;
};

/// @brief 函数信息结构体
using UDFList = std::map<std::wstring, UDFInfo>;

/// @brief UDF注册类
/// @details 通过UDFRegistry类可以注册XLL中的UDF函数
class UDFRegistry {
public:
  /// @brief 获取当前对象 @param name 注册函数的名字(可选) @return 当前对象
  static UDFRegistry& instance(std::wstring name = L"");

  /// @brief 注册UDF函数 @param name 注册函数的名字 @param paramNum 注册函数的参数个数 @return 当前对象
  UDFRegistry* registerFunction(const std::wstring& name, const int& paramNum);

  /// @brief 注册函数
  UDFRegistry* regist();

  /// @brief 取消注册函数
  UDFRegistry* unregist();

  /// @brief 自动注册所有函数
  UDFRegistry* AutoRegist();

  /// @brief 自动取消注册所有函数
  UDFRegistry* AutoUnRegist();

  /// @brief 获取当前对象 @return 当前对象
  UDFRegistry* get_this();

  /// @brief 设置函数在Excel函数中的帮助内容 @param info 函数信息 @return 当前对象
  UDFRegistry* set_info(const wchar_t* info);

  /// @brief 批量设置函数信息 @param info 函数信息字典 @return 当前对象
  /// @note 函数信息格式为: {
  /// @note {udf::name, "value1"},{udf::help, "value2"},
  /// @note {udf::category, "value3"},{udf::type, "value4"},
  /// @note {udf::arguments, "value5"},{udf::args_help, "value6"}
  /// @note {udf::registername, "value7"}
  /// @note }
  UDFRegistry* set_info(std::map<udf::function, std::wstring> info);

  /// @brief 设置注册函数的名称 @param name 函数的注册名称 @return 返回当前对象
  UDFRegistry* set_regsitername(const wchar_t* name);

  /// @brief 设置函数在Excel中的名称 @param name 函数名字 @return 返回当前对象
  UDFRegistry* set_funcname(const wchar_t* name);

  /// @brief 设置函数在Excel中的参数类型 @param text 参数类型占位符如"UUU" @return 返回当前对象
  UDFRegistry* set_typetext(const wchar_t* text);

  /// @brief 设置函数在Excel提示参数 @param text 参数名字(如:"参数1,参数2,参数3"以逗号分割) @return 返回当前对象
  UDFRegistry* set_argstip(const wchar_t* text);

  /// @brief 设置函数在Excel函数的公式所属类别 @param category 所属类别名字 @return 返回当前对象
  UDFRegistry* set_category(const wchar_t* category);

  /// @brief 设置函数在Excel函数中的帮助内容 @param help 帮助文档 @return 返回当前对象
  UDFRegistry* set_funchelp(const wchar_t* help);

  /// @brief 设置函数在Excel中各个参数的帮助内容 @param help 帮助文档(如:"参数1,参数2,参数3"以逗号分割) @return 返回当前对象
  UDFRegistry* set_argshelp(const wchar_t* help);

private:
  std::wstring name;
  UDFList udfs;
};
