#include "xllManager.h"
#include "dll.h"

/// @brief Triggered when opening document @return int
extern "C" __declspec(dllexport) int xlAutoOpen(void) {
    int ret = xll::open();
    UDFRegistry::instance().AutoRegist();
    if (xll::enableRTD) AutoRegisterDll();
    return ret;
}

/// @brief Triggered when closing document @return int
extern "C" __declspec(dllexport) int xlAutoClose(void) {
    UDFRegistry::instance().AutoUnRegist();
    if (xll::enableRTD) DllUnregisterServer();
    return xll::close();
}

/// @brief Triggered when adding xll @return int
extern "C" __declspec(dllexport) int xlAutoAdd(void) {
    return xll::add();
}

/// @brief Triggered when unloading xll @return int
extern "C" __declspec(dllexport) int xlAutoRemove(void) {
    return xll::remove();
}

/// @brief Control memory release of xll functions
extern "C" __declspec(dllexport) void xlAutoFree12(LPXLOPER12 pxFree) {
    if (pxFree->xltype & xltypeStr && pxFree->val.str) {
        delete[] pxFree->val.str;
    } else if (pxFree->xltype == (xltypeMulti | xlbitDLLFree) && pxFree->val.array.lparray) {
        int size = pxFree->val.array.rows * pxFree->val.array.columns;
        LPXLOPER12 p = pxFree->val.array.lparray;
        for (; size-- > 0; p++)
            if (p->xltype & xltypeStr && p->val.str)
                delete[] p->val.str;
        delete[] pxFree->val.array.lparray;
    }
    delete pxFree;
}

///@brief Register xll functions
extern "C" __declspec(dllexport) LPXLOPER12 xlAutoRegister12(LPXLOPER12 pxName) {
    static xloper12 xDLL, xRegId;
    int i;

    xRegId.xltype = xltypeErr;
    xRegId.val.err = xlerrValue;

    // for (i = 0; i < rgFuncsRows; i++)
    // {
    //     if (!lpwstricmp(rgFuncs[i][0].c_str(), pxName->val.str))
    //     {
    //         Excel12f(xlGetName, &xDLL, 0);

    //         Excel12f(xlfRegister, 0, 4,
    //                  (LPXLOPER12)&xDLL,
    //                  (LPXLOPER12)TempStr12(rgFuncs[i][0].c_str()),
    //                  (LPXLOPER12)TempStr12(rgFuncs[i][1].c_str()),
    //                  (LPXLOPER12)TempStr12(rgFuncs[i][2].c_str()));

    //         Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    //         return (LPXLOPER12)&xRegId;
    //     }
    // }
    return (LPXLOPER12)&xRegId;
}

/// @brief xll manager information @param xAction
extern "C" __declspec(dllexport) LPXLOPER12 xlAddInManagerInfo12(LPXLOPER12 xAction) {
    static xloper12 xInfo = {};

    xloper12 xIntAction;
    xloper12 temp = makeXllInt(xltypeInt);
    Excel12(xlCoerce, &xIntAction, 2, xAction, &temp);
    if (xIntAction.val.w == 1) {
        if (xInfo.xltype != xltypeStr) {
            xll::xllName.insert(0, 1, xll::xllName.length());
            xInfo.xltype = xltypeStr;
            xInfo.val.str = xll::xllName.data();
        }
    } else {
        xInfo = makeXllError(xlerrValue);
    }
    return (LPXLOPER12)&xInfo;
}

namespace xll {
bool enableRTD = true;
std::wstring xllName = L"Default";
std::wstring defaultCategory = L"XLL Functions";
XllFunc open = []() { return 1; };
XllFunc close = []() { return 1; };
XllFunc add = []() { return 1; };
XllFunc remove = []() { return 1; };

int MsgBox(const wchar_t* msg, const wchar_t* title) {
    return MessageBoxW(NULL, msg, title, MB_OK);
}

bool alert(const wchar_t* msg) {
    bool ret = false;
    xloper12 xMsg = makeXllStr(makeStr12(msg));
    xloper12 xInt = makeXllInt(2);
    ret = Excel12(xlcAlert, 0, 2, &xMsg, &xInt) == xlretSuccess;
    delete[] xMsg.val.str;
    return ret;
}

bool getCellInfomation(xloper12& cellInfo) {
    xloper12 x;
    bool ret = false;
    if (Excel12(xlfCaller, &x, 0) == xlretSuccess) {
        cellInfo = x;
        ret = true;
        Excel12(xlFree, 0, 1, &x);
    }
    return ret;
}

bool getXLLFullPath(std::wstring& path) {
    xloper12 x;
    bool ret = false;
    if (Excel12(xlGetName, &x, 0) == xlretSuccess) {
        path = unmakeStr12(&x);
        ret = true;
        Excel12(xlFree, 0, 1, &x);
    }
    return ret;
}

bool getExcelHandle(HWND& hwnd) {
    xloper12 x;
    bool ret = false;
    if (Excel12(xlGetHwnd, &x, 0) == xlretSuccess) {
        hwnd = (HWND)((long long)x.val.w);
        ret = true;
        Excel12(xlFree, 0, 1, &x);
    }
    return ret;
}

bool evaluate(std::wstring expression, xllType& value) {
    bool ret = false;
    wchar_t* expr = makeStr12(expression);
    xloper12 v, x = makeXllStr(expr);
    if (Excel12(xlfEvaluate, &v, 1, &x) == xlretSuccess) {
        ret = true;
        value = v;
        Excel12(xlFree, 0, 1, &v);
    }
    delete[] expr;
    return ret;
}
} // namespace xll