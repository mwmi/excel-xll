#include <windows.h>
#include "XLCALL.H"
#include "xlltools.h"
#include "xlcall.cpp"

wchar_t* makeStr12(const wchar_t* ws) {
    int len = wcslen(ws);
    wchar_t* ret = new wchar_t[len + 2];
    ret[0] = (wchar_t)len;
    wcscpy(ret + 1, ws);
    ret[len + 1] = 0;
    return ret;
}

wchar_t* makeStr12(const char* s) {
    int wslen = MultiByteToWideChar(CP_UTF8, 0, s, -1, nullptr, 0);
    wchar_t* ret = new wchar_t[wslen + 1];
    if (ret == nullptr) return nullptr;
    if (MultiByteToWideChar(CP_UTF8, 0, s, -1, ret + 1, wslen) == 0) {
        delete[] ret;
        return nullptr;
    }
    ret[0] = (wchar_t)(wslen - 1);
    ret[wslen] = 0;
    return ret;
}

wchar_t* makeStr12(const xloper12& x) {
    return copyStr12(x.val.str);
}

wchar_t* makeStr12(const std::wstring& ws) {
    int len = ws.size();
    wchar_t* ret = new wchar_t[len + 2];
    ret[0] = (wchar_t)len;
    wcscpy(ret + 1, ws.c_str());
    ret[len + 1] = 0;
    return ret;
}

wchar_t* copyStr(const char* s) {
    int wslen = MultiByteToWideChar(CP_UTF8, 0, s, -1, nullptr, 0);
    wchar_t* ret = new wchar_t[wslen];
    if (ret == nullptr) return nullptr;
    if (MultiByteToWideChar(CP_UTF8, 0, s, -1, ret, wslen) == 0) {
        delete[] ret;
        return nullptr;
    }
    ret[wslen - 1] = 0;
    return ret;
}

wchar_t* copyStr12(const wchar_t* ws) {
    int len = wcslen(ws) + 1;
    wchar_t* ret = new wchar_t[len];
    wcscpy(ret, ws);
    return ret;
}

wchar_t* unmakeStr12(const wchar_t* ws) {
    int len = ws[0];
    wchar_t* ret = new wchar_t[len + 1];
    wcsncpy(ret, ws + 1, len);
    ret[len] = 0;
    return ret;
}

void unmakeStr12(std::wstring& ws, const wchar_t* valstr) {
    wchar_t* t = unmakeStr12(valstr);
    if (t != nullptr) ws = t;
    delete[] t;
}

std::wstring unmakeStr12(LPXLOPER12 xloper) {
    if (!(xloper->xltype & xltypeStr))
        return L"";
    wchar_t* t = unmakeStr12(xloper->val.str);
    std::wstring ret(t);
    delete[] t;
    return ret;
}

xloper12 makeXllStr(wchar_t* ws) {
    xloper12 x;
    x.xltype = xltypeStr;
    x.val.str = ws;
    return x;
}

xloper12 makeXllInt(int i) {
    xloper12 x;
    x.xltype = xltypeInt;
    x.val.w = i;
    return x;
}

xloper12 makeXllError(int err) {
    xloper12 x;
    x.xltype = xltypeErr;
    x.val.err = err;
    return x;
}

xloper12 makeXllNum(double d) {
    xloper12 x;
    x.xltype = xltypeNum;
    x.val.num = d;
    return x;
}

bool xllSerialize(const std::vector<std::vector<std::wstring>>& data, std::wstring& result) {
    // 预先计算所需的总字符数以减少内存重新分配
    size_t total_length = 0;

    // 第一遍：计算所需的总长度
    for (const auto& row : data) {
        if (!row.empty()) {
            // 添加行分隔符（除了第一行）
            if (total_length > 0) total_length += 1; // 对于 | 分隔符

            for (const auto& cell : row) {
                // 添加列分隔符（除了第一列）
                if (total_length > 0 && &cell != &row[0]) total_length += 1; // 对于 , 分隔符

                // 计算单元格内容的长度（包括转义字符）
                for (wchar_t c : cell) {
                    total_length += 1;
                    if (c == L'\\' || c == L',' || c == L'|') {
                        total_length += 1; // 需要额外添加转义字符
                    }
                }
            }
        }
    }

    // 预先分配足够的内存
    result.reserve(total_length);

    // 第二遍：构建字符串
    bool first_row = true;
    for (const auto& row : data) {
        if (row.empty()) continue;

        // 添加行分隔符
        if (!first_row) {
            result += L'|';
        }
        first_row = false;

        bool first_cell = true;
        for (const auto& cell : row) {
            // 添加列分隔符
            if (!first_cell) {
                result += L',';
            }
            first_cell = false;

            // 处理单元格内容
            for (wchar_t c : cell) {
                if (c == L'\\' || c == L',' || c == L'|') {
                    result += L'\\';
                }
                result += c;
            }
        }
    }

    return !result.empty();
}

bool xllDeserialize(const std::wstring& str, std::vector<std::vector<std::wstring>>& result) {
    std::vector<std::wstring> currentRow;
    std::wstring currentValue;

    // 预先估计容量以减少重新分配
    currentValue.reserve(64);  // 假设平均字段长度

    bool escaping = false;

    for (wchar_t c : str) {
        if (escaping) {
            currentValue += c;
            escaping = false;
        } else if (c == L'\\') {
            escaping = true;
        } else if (c == L',') {
            currentRow.emplace_back(std::move(currentValue));
            currentValue.reserve(64);  // 保持预留容量
        } else if (c == L'|') {
            currentRow.emplace_back(std::move(currentValue));
            result.emplace_back(std::move(currentRow));
            currentValue.reserve(64);  // 保持预留容量
        } else {
            currentValue += c;
        }
    }

    // 处理最后一个值
    if (currentValue.empty()) {
        int size = str.size();
        if ((size >= 2) && (str[size - 1] == L',') && (str[size - 2] != L'\\' || (size > 2 && str[size - 3] == L'\\'))) {
            currentRow.emplace_back(L"");
        }
    } else {
        currentRow.emplace_back(std::move(currentValue));
    }
    if (!currentRow.empty()) {
        result.emplace_back(std::move(currentRow));
    }
    if (result.empty()) return false;
    return true;
}