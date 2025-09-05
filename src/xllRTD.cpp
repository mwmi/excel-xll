#include "xllRTD.h"

RTDRegister& RTDRegister::instance() {
    static RTDRegister instance;
    return instance;
}

void RTDRegister::registerRTDFunction(const std::wstring& name, RtdFun fun, const wchar_t * default_value, bool is_async) {
    _async_functions[name] = fun;
    _default_values[name] = default_value;
    _is_async[name] = is_async;

}

void RTDRegister::registerRTDFunction(const std::wstring& name, RtdFun fun, bool is_async) {
    _async_functions[name] = fun;
    _default_values[name] = L"";
    _is_async[name] = is_async;
}

int RTDRegister::runAsyncFunction(const std::wstring& name, xllptrlist& args, Topic* topic) {
    return _async_functions[name](std::move(args), topic);
}

bool RTDRegister::isFunctionRegistered(const std::wstring& name) {
    return _async_functions.find(name) != _async_functions.end();
}

bool RTDRegister::getDefaultValue(const std::wstring& name, std::wstring& default_value) {
    if (_default_values.find(name) != _default_values.end()) {
        default_value = _default_values[name];
        return true;
    }
    return false;
}

bool RTDRegister::isFunctionAsync(const std::wstring& name) {
    if (_is_async.find(name) != _is_async.end()) {
        return _is_async[name];
    }
    return false;
}

RtdFun& RTDRegister::getFunction(const std::wstring& name) {
    return _async_functions[name];
}

int registerRTDTask(Topic* topic) {
    RTDRegister& rtd = RTDRegister::instance();
    size_t count = topic->getArgCount();
    if (count < 1) return -1;
    const std::wstring funcName = topic->getArg(0);
    if (!rtd.isFunctionRegistered(funcName)) return -1;
    bool is_async = rtd.isFunctionAsync(funcName);
    std::wstring default_value;
    if (rtd.getDefaultValue(funcName, default_value)) {
        topic->setDefaultValue(default_value);
    } else {
        topic->setDefaultValue(L"");
    }
    topic->setTask([&rtd, count, funcName](Topic* topic) {
        xllptrlist args;
        for (int i = 1; i < count; i++) {
            xllType arg = topic->getArg(i);
            args.push_back(std::make_unique<xllType>(*(arg.deserialize())));
        }
        return rtd.runAsyncFunction(funcName, args, topic);
    }, is_async);
    return 0;
}