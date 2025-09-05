#pragma once

#include "xllType.h"
#include <atomic>
#include <functional>
#include <mutex>
#include <string>
#include <vector>

// Forward declaration
class Topic;

// Type aliases
using Task = std::function<int(Topic* topic)>;
using StringArray = std::vector<std::wstring>;
using StringMatrix = std::vector<StringArray>;

// VARIANT creation function declarations
VARIANT createVariant(int value);
VARIANT createVariant(const std::wstring& value);

/**
 * @class Topic
 * @brief RTD topic class for managing real-time data topics
 *
 * This class encapsulates the core functionality of RTD topics, including:
 * - Topic parameter management
 * - Value storage and updates
 * - Asynchronous task execution
 * - COM object interaction
 */
class Topic {
public:
  // Constructors and destructor
  Topic() = default;
  ~Topic();
  Topic(long id, SAFEARRAY** Strings);
  Topic(long id, SAFEARRAY** Strings, std::wstring defaultValue);

  // Disable copy constructor and assignment operator (contains Windows handles)
  Topic(const Topic&) = delete;
  Topic& operator=(const Topic&) = delete;

  // Basic information access
  long getID() const;
  std::wstring getArg(size_t index) const;
  size_t getArgCount() const;

  // Default value management
  bool hasDefaultValue() const;
  Topic* setDefaultValue(std::wstring value);
  std::wstring getDefaultValue() const;

  // Value management
  bool hasValue() const;
  Topic* setValue(const std::wstring& value);
  Topic* setValue(xllType& x);
  std::wstring getValue() const;
  bool hasChanged() const;
  Topic* update(SAFEARRAY** parrayOut, int i);

  // Task management
  Topic* setAsync(bool isAsync);
  Topic* setTask(Task task, bool is_async = false, int run_count = 1);
  bool isTaskRunning() const;
  Topic* stopTask();
  bool runTask();

private:
  // Member variables
  long topic_id = 0;                   // Topic ID
  int task_run_count = 1;              // Task execution count
  StringArray args;                    // Parameter array
  Task task = nullptr;                 // Task function
  bool isAsync = false;                // Whether to execute asynchronously
  HANDLE async_handle = nullptr;       // Async thread handle
  std::atomic<bool> is_runing = false; // Running status flag
  std::wstring default_value;          // Default value
  std::wstring old_value;              // Old value
  std::wstring value;                  // Current value
  mutable std::mutex mutex_value;      // Mutex lock
  mutable std::mutex mutex_task;       // Mutex lock

  // Private utility functions
  void cleanup();
};
