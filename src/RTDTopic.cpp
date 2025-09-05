#include "RTDTopic.h"

// VARIANT 创建函数实现
VARIANT createVariant(int value) {
    VARIANT variant;
    VariantInit(&variant);
    variant.vt = VT_I4;
    variant.lVal = value;
    return variant;
}

VARIANT createVariant(const std::wstring& value) {
    VARIANT variant;
    VariantInit(&variant);
    variant.vt = VT_BSTR;
    variant.bstrVal = SysAllocString(value.c_str());
    return variant;
}

// Topic 类析构函数
Topic::~Topic() {
    stopTask();
}

// Topic 类构造函数实现
Topic::Topic(long id, SAFEARRAY** Strings) {
    this->topic_id = id;
    if (Strings != nullptr) {
        long lElements = (*Strings)->rgsabound[0].cElements;
        this->args.resize(lElements, L"");
        for (long i = 0; i < lElements; ++i) {
            VARIANT var;
            SafeArrayGetElement(*Strings, &i, &var);
            if (var.vt == VT_BSTR) {
                this->args[i] = var.bstrVal;
            }
            VariantClear(&var);
        }
    }
}

Topic::Topic(long id, SAFEARRAY** Strings, std::wstring defaultValue) : Topic(id, Strings) {
    setDefaultValue(defaultValue);
}

// 私有辅助函数实现
void Topic::cleanup() {
    stopTask();
}

// 成员函数实现
long Topic::getID() const {
    return topic_id;
}

std::wstring Topic::getArg(size_t index) const {
    if (index < args.size()) {
        return args[index];
    }
    return L"";
}

size_t Topic::getArgCount() const {
    return args.size();
}

bool Topic::hasDefaultValue() const {
    std::lock_guard<std::mutex> lock(mutex_value);
    return !default_value.empty();
}

Topic* Topic::setDefaultValue(std::wstring value) {
    std::lock_guard<std::mutex> lock(mutex_value);
    this->default_value = value;
    return this;
}

std::wstring Topic::getDefaultValue() const {
    std::lock_guard<std::mutex> lock(mutex_value);
    return default_value;
}

bool Topic::hasValue() const {
    std::lock_guard<std::mutex> lock(mutex_value);
    return !value.empty();
}

Topic* Topic::setValue(const std::wstring& value) {
    std::lock_guard<std::mutex> lock(mutex_value);
    this->value = value;
    return this;
}

Topic* Topic::setValue(xllType& x) {
    std::lock_guard<std::mutex> lock(mutex_value);
    this->value = x.serialize()->get_str();
    return this;
}

std::wstring Topic::getValue() const {
    std::lock_guard<std::mutex> lock(mutex_value);
    return value;
}

bool Topic::hasChanged() const {
    std::lock_guard<std::mutex> lock(mutex_value);
    return old_value != value;
}

Topic* Topic::update(SAFEARRAY** parrayOut, int i) {
    VARIANT id = createVariant(this->topic_id);
    VARIANT val;
    LONG index[2] = {0, i};
    if (hasValue()) {
        val = createVariant(getValue());
    } else {
        if (default_value.empty()) {
            default_value = L"没有初始值";
        }
        val = createVariant(getDefaultValue());
    }
    SafeArrayPutElement(*parrayOut, index, &id);
    index[0] = 1;
    SafeArrayPutElement(*parrayOut, index, &val);
    VariantClear(&id);
    VariantClear(&val);
    old_value = value;
    return this;
}

Topic* Topic::setAsync(bool isAsync) {
    this->isAsync = isAsync;
    return this;
}

Topic* Topic::setTask(Task task, bool is_async, int run_count) {
    this->task = task;
    this->isAsync = is_async;
    this->task_run_count = run_count;
    return this;
}

bool Topic::isTaskRunning() const {
    return is_runing.load();
}

Topic* Topic::stopTask() {
    std::lock_guard<std::mutex> lock(this->mutex_task);
    if (is_runing.load() && isAsync && async_handle != nullptr) {
        is_runing = false;
        TerminateThread(async_handle, 0);
        CloseHandle(this->async_handle);
        async_handle = nullptr;
    }
    return this;
}

bool Topic::runTask() {
    if (task != nullptr && task_run_count > 0) {
        if (is_runing.load()) {
            return true;
        }
        this->is_runing = true;
        if (isAsync) {
            async_handle = CreateThread(nullptr, 0, [](LPVOID param) -> DWORD {
                Topic* self = static_cast<Topic*>(param);
                int ret = self->task(self);
                self->task_run_count--;
                self->is_runing = false;
                std::lock_guard<std::mutex> lock(self->mutex_task);
                if (self->async_handle) CloseHandle(self->async_handle);
                self->async_handle = nullptr;
                return ret; }, this, 0, nullptr);
        } else {
            this->task(this);
            this->task_run_count--;
            this->is_runing = false;
        }
        return true;
    } else {
        return false;
    }
}