/**
 * @file functions.cpp
 * @brief xll function implementation
 * @author mwmi
 * @copyright Copyright (c) 2025 by mwmi, All rights reserved.
 */
#include "xllManager.h"

UDF(HelloWorld, L"Test text") {
    xllType result;
    result = L"Hello World";
    return result.get_return();
}

UDF(Add, (
    {udf::help, L"Test addition"},
    // { udf::category, L"Category"}
    ), Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    if (a_.is_num() && b_.is_num()) {
        result = a_.get_num() + b_.get_num();
    } else {
        result = L"Type error!";
    }
    return result.get_return();
}

UDF(Concat2, L"Test string concatenation", Param a, Param b) {
    xllType result;
    xllType a_ = a;
    xllType b_ = b;
    std::wstring str = a_.get_str() + b_.get_str();
    result = str;
    return result.get_return();
}

UDF(MySum, L"Test cell reference", Param a) {
    xllType result;
    xllType a_ = a;
    double sum = 0;
    if (a_.is_array()) {
        // Use iterator to traverse each element in the array
        sum = 0;
        for (auto i : a_) {
            if (i->is_num())
                sum += i->get_num();
        }

        // Use index method 1 to traverse the array
        sum = 0;
        for (int i = 0; i < a_.size(); ++i) {
            if ((a_)[i]->is_num())
                sum += (a_)[i]->get_num();
        }

        // Use index method 2 to traverse the array
        sum = 0;
        for (int i = 0; i < a_.size(); ++i) {
            if (a_.at(i)->is_num())
                sum += a_.at(i)->get_num();
        }
    }
    result = sum;
    return result.get_return();
}

UDF(MyConcat, L"Test text array", Param a) {
    xllType result;
    xllType a_ = a;
    std::wstring ret = L"";
    if (a_.is_array()) {
        // Traverse each element in the array and concatenate all string-type cells
        for (auto i : a_) {
            if (i->is_str())
                ret += i->get_str();
        }
    }
    result = ret;
    return result.get_return();
}

UDF(RetArray, L"Test return array") {
    xllType result;
    xllType a = L"String1";
    xllType b = 20;
    xllType c = 30;
    xllType d = 40;
    xllType e = L"String2";

    // Combine variables a, b, c, d into a 2D array structure and assign to result
    result = {{a,b},{c,d}};
    // result.push_back(e);
    // result.push_back(L"Really seriously");
    // result.push_back(12331);
    // result = 10;
    // result.push_back(20);

    // Call the get_return method of xllType to return the result
    return result.get_return();
}

// Test Excel built-in function call
UDF(Test, L"Test built-in function") {
    xllType ret;
    ret = RtdServer_DllPath;
    // xll::callExcelFunction(ret, xlfAbs, -100); // Pass
    // xll::callExcelFunction(ret, xlfSum, a); // Pass
    // xll::callExcelFunction(ret, xlfMin, 10, 20, 30, 40, 50, 60); // Pass
    // xll::callExcelFunction(ret, xlfLeft, L"asdasdasd", 2); // Pass
    return ret.get_return();
}

// RTD Hello World Call: =RTDHelloWorld()
RTD(RTDHelloWorld, L"RTD Hello World", ([](xllptrlist args, Topic* topic) {

    // Define your RTD function logic
    topic->setValue(L"Hello World");
    return 0;

}, L"Loading...")) {

    xllType ret;
    // Call RTD function and pass parameters
    CALLRTD(ret);
    // Return result
    return ret.get_return();

}

// RTD Display Clock (Asynchronous Call) Call: =RTDClock()
RTD(RTDClock, L"Display Clock", ([](xllptrlist args, Topic* topic) {

    SYSTEMTIME st;
    while (true) {
        GetLocalTime(&st);
        wchar_t buffer[100];
        swprintf(buffer, 100, L"【%d】🕒 %04d-%02d-%02d %02d:%02d:%02d", topic->getID(), st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
        topic->setValue(buffer);
        Sleep(500);
    }
    return 0;

}, L"Ready to display", true)) {

    xllType ret;
    CALLRTD(ret);
    return ret.get_return();

}

// Test RTD Array Call: =RTDArray()
RTD(RTDArray, L"Return Array", ([](xllptrlist args, Topic* topic) {
    xllType a = 10.123123;
    xllType b = L"Ten";
    xllType c = 20;
    xllType d = L"Twenty";
    xllType ret({{a,b},{c,d}});
    topic->setValue(ret);
    return 0;
}, L"Preparing data...")) {
    xllType ret;
    CALLRTD(ret);
    return ret.get_return();
}

// Test RTD Parameter Reference Call: =RTDParam(cell reference)
RTD(RTDParam, L"Test passing cell reference and return cell content", ([](xllptrlist args, Topic* topic) {

    xllType p;
    if (args.size() > 0) {
        p = *args[0].get();
    }
    topic->setValue(p);
    return 0;

}, L"Start testing"), Param a) {

    xllType ret;
    CALLRTD(ret, a);
    return ret.get_return();

}

// Configuration settings
SET() {
    xll::xllName = L"XLL Name Setting";

    // Set default XLL category to "Custom Functions".
    // xll::defaultCategory = L"Custom Functions";

    // Set whether to enable RTD, enabled by default
    // xll::enableRTD = false;

    xll::open = []() {
        // Set function information
        // UDFCONFIG(HelloWorld)->set_funchelp(L"Hello World!!!!");
        // xll::alert(L"Welcome to use XLL Loader");
        return 1;
    };

    xll::close = []() {
        return 1;
    };

    return 0;
}