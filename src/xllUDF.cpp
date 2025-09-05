#include <windows.h>
#include "XLCALL.H"
#include "xllManager.h"
#include "xllTools.h"
#include "xllUDF.h"

UDFRegistry& UDFRegistry::instance(std::wstring name) {
    static UDFRegistry registry;
    registry.name = name;
    return registry;
}

UDFRegistry* UDFRegistry::registerFunction(const std::wstring& name, const int& paramNum) {
    this->name = name;
    UDFInfo info;
    info.paramNum = paramNum;
    udfs[name] = info;
    return this;
}

UDFRegistry* UDFRegistry::get_this() { return this; }

UDFRegistry* UDFRegistry::set_info(const wchar_t* info) {
    return this->set_funchelp(info);
}

UDFRegistry* UDFRegistry::set_info(std::map<udf::function, std::wstring> info) {
    for (auto& p : info) {
        switch (p.first) {
        case udf::category:
        this->set_category(p.second.c_str());
        break;
        case udf::help:
        this->set_funchelp(p.second.c_str());
        break;
        case udf::args_help:
        this->set_argshelp(p.second.c_str());
        break;
        case udf::arguments:
        this->set_argstip(p.second.c_str());
        break;
        case udf::name:
        this->set_funcname(p.second.c_str());
        break;
        case udf::type:
        this->set_typetext(p.second.c_str());
        break;
        case udf::registername:
        this->set_regsitername(p.second.c_str());
        break;
        default:
        break;
        }
    }
    return this;
}

UDFRegistry* UDFRegistry::regist() {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];

    XLOPER12 xDLL;
    Excel12(xlGetName, &xDLL, 0);

    int _param_nums = info.paramNum + 1;

    if (info.register_name == nullptr) {
        info.register_name = makeStr12(name);
    }
    if (info.type_text == nullptr) {
        wchar_t* _function_paramater = new wchar_t[_param_nums + 2]();
        _function_paramater[0] = (wchar_t)_param_nums;
        for (int i = 1; i < _param_nums + 1; i++) {
            _function_paramater[i] = L'U';
        }
        info.type_text = _function_paramater;
    }
    if (info.function_name == nullptr) {
        info.function_name = makeStr12(name);
    }
    if (info.argument_text == nullptr) {
        if (_param_nums > 1) {
            std::wstring s = L"参数1";
            for (int i = 2; i < _param_nums; i++)
                s += L",参数" + std::to_wstring(i);
            info.argument_text = makeStr12(s);
        } else {
            info.argument_text = makeStr12(L"");
        }
    }

    if (info.category == nullptr) {
        info.category = makeStr12(xll::defaultCategory);
    }

    if (info.function_help == nullptr) {
        info.function_help = makeStr12(L"");
    }
    if (info.argument_help == nullptr) {
        info.argument_help = makeStr12(L"");
    }

    XLOPER12 _register_function_name = makeXllStr(info.register_name);
    XLOPER12 _type_text = makeXllStr(info.type_text);
    XLOPER12 _function_name = makeXllStr(info.function_name);
    XLOPER12 _argument_text = makeXllStr(info.argument_text);
    XLOPER12 _category = makeXllStr(info.category);
    XLOPER12 _function_help = makeXllStr(info.function_help);
    XLOPER12 _argument_help = makeXllStr(info.argument_help);
    XLOPER12 _t1 = makeXllStr((wchar_t*)L"\0011");
    XLOPER12 _t2 = makeXllStr((wchar_t*)L"\000");

    Excel12(xlfRegister, 0, 10, &xDLL,
        &_register_function_name,
        &_type_text,
        &_function_name,
        &_argument_text,
        &_t1,
        &_category,
        &_t2,
        &_t2,
        &_function_help,
        &_argument_help);

    Excel12(xlFree, 0, 1, &xDLL);
    return this;
}

UDFRegistry* UDFRegistry::unregist() {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];

    if (info.register_name) {
        XLOPER12 t;
        t.xltype = xltypeStr;
        t.val.str = info.register_name;
        Excel12(xlfSetName, 0, 1, &t);
        delete[] info.register_name;
        info.register_name = nullptr;
    }
    if (info.type_text) {
        delete[] info.type_text;
        info.type_text = nullptr;
    }
    if (info.function_name) {
        delete[] info.function_name;
        info.function_name = nullptr;
    }
    if (info.argument_text) {
        delete[] info.argument_text;
        info.argument_text = nullptr;
    }
    if (info.category) {
        delete[] info.category;
        info.category = nullptr;
    }
    if (info.function_help) {
        delete[] info.function_help;
        info.function_help = nullptr;
    }
    if (info.argument_help) {
        delete[] info.argument_help;
        info.argument_help = nullptr;
    }
    return this;
}

UDFRegistry* UDFRegistry::AutoRegist() {
    for (auto& p : this->udfs) {
        this->name = p.first;
        this->regist();
    }
    return this;
}

UDFRegistry* UDFRegistry::AutoUnRegist() {
    for (auto& p : this->udfs) {
        this->name = p.first;
        this->unregist();
    }
    return this;
}

UDFRegistry* UDFRegistry::set_regsitername(const wchar_t* name) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.register_name) {
        delete[] info.register_name;
        info.register_name = nullptr;
    }
    info.register_name = makeStr12(name);
    return this;
}

UDFRegistry* UDFRegistry::set_funcname(const wchar_t* name) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.function_name) {
        delete[] info.function_name;
        info.function_name = nullptr;
    }
    info.function_name = makeStr12(name);
    return this;
}

UDFRegistry* UDFRegistry::set_typetext(const wchar_t* text) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.type_text) {
        delete[] info.type_text;
        info.type_text = nullptr;
    }
    info.type_text = makeStr12(text);
    return this;
}

UDFRegistry* UDFRegistry::set_argstip(const wchar_t* text) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.argument_text) {
        delete[] info.argument_text;
        info.argument_text = nullptr;
    }
    info.argument_text = makeStr12(text);
    return this;
}

UDFRegistry* UDFRegistry::set_category(const wchar_t* category) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.category) {
        delete[] info.category;
        info.category = nullptr;
    }
    info.category = makeStr12(category);
    return this;
}

UDFRegistry* UDFRegistry::set_funchelp(const wchar_t* help) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.function_help) {
        delete[] info.function_help;
        info.function_help = nullptr;
    }
    // 添加空格，避免函数提示被截断
    info.function_help = makeStr12(std::wstring(help) + L' ');
    return this;
}

UDFRegistry* UDFRegistry::set_argshelp(const wchar_t* help) {
    if (this->udfs.find(this->name) == this->udfs.end()) return this;
    UDFInfo& info = this->udfs[this->name];
    if (info.argument_help) {
        delete[] info.argument_help;
        info.argument_help = nullptr;
    }
    info.argument_help = makeStr12(help);
    return this;
}
