/**
 * @file functions.cpp
 * @brief xll函数实现
 * @author mwmi
 * @copyright Copyright (c) 2025 by mwmi, All rights reserved.
 */
#include "xllManager.h"

UDF(HelloWorld, L"测试文本") {
    xllType result;
    result = L"你好世界";
    return result.get_return();
}

UDF(Add, (
    {udf::help, L"测试加法"},
    // { udf::category, L"分类"}
    ), Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    if (a_.is_num() && b_.is_num()) {
        result = a_.get_num() + b_.get_num();
    } else {
        result = L"类型错误！";
    }
    return result.get_return();
}

UDF(Concat2, L"测试拼接字符串", Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    std::wstring str = a_.get_str() + b_.get_str();
    result = str;
    return result.get_return();
}

UDF(MySum, L"测试单元格引用", Param a) {
    xllType result;
    xllType a_ = a;
    double sum = 0;
    if (a_.is_array()) {
        // 使用迭代器遍历数组中的每个元素
        sum = 0;
        for (auto i : a_) {
            if (i->is_num())
                sum += i->get_num();
        }

        // 使用索引方式1遍历数组
        sum = 0;
        for (int i = 0; i < a_.size(); ++i) {
            if ((a_)[i]->is_num())
                sum += (a_)[i]->get_num();
        }

        // 使用索引方式2遍历数组
        sum = 0;
        for (int i = 0; i < a_.size(); ++i) {
            if (a_.at(i)->is_num())
                sum += a_.at(i)->get_num();
        }
    }
    result = sum;
    return result.get_return();
}

UDF(MyConcat, L"测试文本数组", Param a) {
    xllType result;
    xllType a_ = a;
    std::wstring ret = L"";
    if (a_.is_array()) {
        // 遍历数组中的每个元素，连接所有字符串类型的单元格
        for (auto i : a_) {
            if (i->is_str())
                ret += i->get_str();
        }
    }
    result = ret;
    return result.get_return();
}

UDF(RetArray, L"测试返回数组") {
    xllType result;
    xllType a = L"字符串1";
    xllType b = 20;
    xllType c = 30;
    xllType d = 40;
    xllType e = L"字符串2";

    // 将变量 a, b, c, d 组合成一个二维数组结构并赋值给 result
    result = {{a,b},{c,d}};
    // result.push_back(e);
    // result.push_back(L"你啊实打实的");
    // result.push_back(12331);
    // result = 10;
    // result.push_back(20);

    // 调用 xllType 的 get_return 方法返回结果
    return result.get_return();
}

// 测试Excel内置函数调用
UDF(Test, L"测试内置函数") {
    xllType ret;
    ret = RtdServer_DllPath;
    // xll::callExcelFunction(ret, xlfAbs, -100); // 通过
    // xll::callExcelFunction(ret, xlfSum, a); // 通过
    // xll::callExcelFunction(ret, xlfMin, 10, 20, 30, 40, 50, 60); // 通过
    // xll::callExcelFunction(ret, xlfLeft, L"asdasdasd", 2); // 通过
    return ret.get_return();
}

// RTD你好世界 调用: =RTDHelloWorld()
RTD(RTDHelloWorld, L"RTD你好世界", ([](xllptrlist args, Topic* topic) {

    // 定义你的RTD函数逻辑
    topic->setValue(L"Hello World");
    return 0;

}, L"等待加载中...")) {

    xllType ret;
    // 调用RTD函数，并传入参数
    CALLRTD(ret);
    // 返回结果
    return ret.get_return();

}

// RTD显示时钟(异步调用) 调用: =RTDClock()
RTD(RTDClock, L"显示时钟", ([](xllptrlist args, Topic* topic) {

    SYSTEMTIME st;
    while (true) {
        GetLocalTime(&st);
        wchar_t buffer[100];
        swprintf(buffer, 100, L"【%d】🕒 %04d-%02d-%02d %02d:%02d:%02d", topic->getID(), st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
        topic->setValue(buffer);
        Sleep(500);
    }
    return 0;

}, L"准备显示", true)) {

    xllType ret;
    CALLRTD(ret);
    return ret.get_return();

}

// 测试RTD数组 调用: =RTDArray()
RTD(RTDArray, L"返回数组", ([](xllptrlist args, Topic* topic) {
    xllType a = 10.123123;
    xllType b = L"十";
    xllType c = 20;
    xllType d = L"二十";
    xllType ret({{a,b},{c,d}});
    topic->setValue(ret);
    return 0;
}, L"准备数据中...")) {
    xllType ret;
    CALLRTD(ret);
    return ret.get_return();
}

// 测试RTD参数引用 调用: =RTDParam(单元格引用)
RTD(RTDParam, L"测试传入单元格引用，并返回单元格内容", ([](xllptrlist args, Topic* topic) {

    xllType p;
    if (args.size() > 0) {
        p = *args[0].get();
    }
    topic->setValue(p);
    return 0;

}, L"开始测试"), Param a) {

    xllType ret;
    CALLRTD(ret, a);
    return ret.get_return();

}

// 配置设置
SET() {
    xll::xllName = L"xll名称设置";

    // 设置默认的XLL默认类别为“自定义功能”。
    // xll::defaultCategory = L"自定义功能";

    // 设置是否启用RTD, 默认启用
    // xll::enableRTD = false;

    xll::open = []() {
        // 设置函数信息
        // UDFCONFIG(HelloWorld)->set_funchelp(L"你好世界!!!!");
        // xll::alert(L"欢迎使用XLL加载器");
        return 1;
    };

    xll::close = []() {
        return 1;
    };

    return 0;
}