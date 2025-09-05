#pragma once

#include "xllType.h"
#include <atomic>
#include <functional>
#include <mutex>
#include <string>
#include <vector>

// 前向声明
class Topic;

// 类型别名
using Task = std::function<int(Topic* topic)>;
using StringArray = std::vector<std::wstring>;
using StringMatrix = std::vector<StringArray>;

// VARIANT 创建函数声明
VARIANT createVariant(int value);
VARIANT createVariant(const std::wstring& value);

/**
 * @class Topic
 * @brief RTD 主题类，用于管理实时数据主题
 *
 * 该类封装了 RTD 主题的核心功能，包括：
 * - 主题参数管理
 * - 值的存储和更新
 * - 异步任务执行
 * - COM 对象交互
 */
class Topic {
public:
  // 构造函数和析构函数
  Topic() = default;
  ~Topic();
  Topic(long id, SAFEARRAY** Strings);
  Topic(long id, SAFEARRAY** Strings, std::wstring defaultValue);

  // 禁用拷贝构造和赋值运算符（因为包含 Windows 句柄）
  Topic(const Topic&) = delete;
  Topic& operator=(const Topic&) = delete;

  // 基本信息访问
  long getID() const;
  std::wstring getArg(size_t index) const;
  size_t getArgCount() const;

  // 默认值管理
  bool hasDefaultValue() const;
  Topic* setDefaultValue(std::wstring value);
  std::wstring getDefaultValue() const;

  // 值管理
  bool hasValue() const;
  Topic* setValue(const std::wstring& value);
  Topic* setValue(xllType& x);
  std::wstring getValue() const;
  bool hasChanged() const;
  Topic* update(SAFEARRAY** parrayOut, int i);

  // 任务管理
  Topic* setAsync(bool isAsync);
  Topic* setTask(Task task, bool is_async = false, int run_count = 1);
  bool isTaskRunning() const;
  Topic* stopTask();
  bool runTask();

private:
  // 成员变量
  long topic_id = 0;                   // 主题ID
  int task_run_count = 1;              // 任务运行次数
  StringArray args;                    // 参数数组
  Task task = nullptr;                 // 任务函数
  bool isAsync = false;                // 是否异步执行
  HANDLE async_handle = nullptr;       // 异步线程句柄
  std::atomic<bool> is_runing = false; // 运行状态标志
  std::wstring default_value;          // 默认值
  std::wstring old_value;              // 旧值
  std::wstring value;                  // 当前值
  mutable std::mutex mutex_value;      // 互斥锁
  mutable std::mutex mutex_task;       // 互斥锁

  // 私有辅助函数
  void cleanup();
};
