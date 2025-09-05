#include "xllType.h"
#include "xllTools.h"
#include "xllManager.h"

xllType* xllType::init() {
    this->xltype = xltypeNil;
    this->val.str = nullptr;
    this->val.array.lparray = nullptr;
    return this;
}

xllType* xllType::init(const xloper12& x) {
    this->xltype = x.xltype;
    this->val = x.val;
    return this->load();
}

xllType* xllType::destory() {
    this->num = 0;
    this->str.clear();
    this->array.clear();
    this->optr.reset(nullptr);
    return this;
}

xllType* xllType::load() {
    if (this->is_array()) {
        if (!check_ref() || !this->load_ref(xltypeMulti)) this->set_err(xlerrRef);
        return this;
    }
    // Numeric judgment must be placed before string judgment here, because numbers can be converted to strings during reading
    this->load_ref(xltypeNum);
    if (this->is_num()) {
        if (this->is_sref()) {
            this->num = this->optr->val.num;
        } else {
            this->num = this->val.num;
        }
        this->xltype = xltypeNum;
    }
    this->load_ref(xltypeStr);
    if (this->is_str()) {
        if (this->is_sref()) {
            this->str = unmakeStr12(this->optr->val.str);
        } else {
            this->str = unmakeStr12(this->val.str);
        }
        this->xltype = xltypeStr;
    }
    return this;
}

xllType* xllType::copy(const xllType& other) {
    this->xltype = other.xltype;
    this->val = other.val;
    this->num = other.num;
    this->str = other.str;
    this->rows = other.rows;
    this->cols = other.cols;
    for (auto& x : other.array) {
        if (x) this->array.emplace_back(std::make_unique<xllType>(*x));
    }
    return this;
}

xllType* xllType::set_err(int err) {
    this->destory();
    this->xltype = xltypeErr;
    this->val.err = err;
    return this;
}

xllType* xllType::set_num(double num) {
    this->destory();
    this->xltype = xltypeNum;
    this->num = num;
    return this;
}

xllType* xllType::set_str(const wchar_t* ws) {
    this->destory();
    this->xltype = xltypeStr;
    this->str = ws;
    return this;
}

xllType* xllType::set_str(const std::wstring& ws) {
    this->destory();
    this->xltype = xltypeStr;
    this->str = ws;
    return this;
}

xllType* xllType::set_list(const xlllist& l) {
    if (l.empty()) return this;
    this->destory();
    for (auto& x : l) {
        this->array.emplace_back(std::make_unique<xllType>(x));
    }
    this->rows = 1;
    this->cols = this->array.size();
    this->xltype = xltypeMulti;
    return this;
}

xllType* xllType::set_matrix(const xllmartix& m) {
    if (m.empty()) return this;
    this->destory();
    for (auto& x : m) {
        for (auto& y : x) {
            this->array.emplace_back(std::make_unique<xllType>(y));
        }
    }
    this->rows = m.size();
    this->cols = m[0].size();
    this->xltype = xltypeMulti;
    return this;
}

double xllType::get_num() const {
    return this->num;
}

std::wstring xllType::get_str() const {
    // if (this->str.empty()) return L"";
    return this->str;
}

const wchar_t* xllType::get_c_str() const {
    return this->str.c_str();
}

int xllType::get_last_err() const {
    return this->error_code;
}

bool xllType::load_ref(DWORD type) {
    if (!this->is_sref()) return false;
    xloper12 x, t = makeXllInt(type);
    int r = Excel12(xlCoerce, &x, 2, this, &t);
    if (r == xlretSuccess) {
        this->optr = std::make_unique<xloper12>(x);
        if (xltypeMulti == type) {
            this->rows = this->optr->val.array.rows;
            this->cols = this->optr->val.array.columns;
            int n = this->rows * this->cols;
            if (!this->array.empty()) this->array.clear();
            if (!this->optr->val.array.lparray) {
                Excel12(xlFree, 0, 1, &x);
                return false;
            }
            for (int i = 0; i < n; i++) {
                this->array.emplace_back(std::make_unique<xllType>(this->optr->val.array.lparray + i));
            }
        }
        Excel12(xlFree, 0, 1, &x);
    } else {
        this->error_code = r;
    }
    return r == xlretSuccess;
}

bool xllType::check_ref() {
    xloper12 info;
    if (xll::getCellInfomation(info)) {
        int row = info.val.sref.ref.rwFirst;
        int col = info.val.sref.ref.colFirst;
        if (this->val.sref.ref.rwFirst <= row && row <= this->val.sref.ref.rwLast) {
            if (this->val.sref.ref.colFirst <= col && col <= this->val.sref.ref.colLast) {
                return false;
            }
        }
    }
    return true;
}


xllType::xllType() {
    this->init();
}

xllType::~xllType() {
    this->destory();
}

xllType::xllType(const xllType& other) {
    this->copy(other);
}

xllType::xllType(xllType&& other) noexcept {
    this->xltype = other.xltype;
    this->val = other.val;
    this->num = other.num;
    this->str = other.str;
    this->rows = other.rows;
    this->cols = other.cols;
    this->optr = std::move(other.optr);
    other.destory();
}

xllType::xllType(const xloper12* px) {
    this->xltype = px->xltype;
    this->val = px->val;
    this->load();
}

xllType::xllType(const xloper12& x) {
    this->xltype = x.xltype;
    this->val = x.val;
    this->load();
}

xllType::xllType(double num) {
    this->init()->xltype = xltypeNum;
    this->num = num;
}

xllType::xllType(const char* str) {
    this->init()->xltype = xltypeStr;
    const wchar_t* ws = copyStr(str);
    if (ws) {
        this->str = ws;
        delete[] ws;
    }
}

xllType::xllType(const wchar_t* ws) {
    this->init()->xltype = xltypeStr;
    this->str = ws;
}

xllType::xllType(const std::wstring& str) {
    this->init()->xltype = xltypeStr;
    this->str = str;
}

xllType::xllType(const xlllist& l) {
    this->init()->xltype = xltypeMulti;
    for (auto& x : l) {
        this->array.emplace_back(std::make_unique<xllType>(x));
    }
    this->rows = 1;
    this->cols = this->array.size();
}

xllType::xllType(const xllmartix& m) {
    this->init()->xltype = xltypeMulti;
    for (auto& x : m) {
        for (auto& y : x) {
            this->array.emplace_back(std::make_unique<xllType>(y));
        }
    }
    this->rows = m.size();
    this->cols = m[0].size();
}

xllType& xllType::operator=(const xllType& other) {
    return *(this->destory()->copy(other));
}

xllType& xllType::operator=(const xloper12& x) {
    return *(this->destory()->init(x));
}

xllType& xllType::operator=(double num) {
    return *(this->set_num(num));
}

xllType& xllType::operator=(const wchar_t* ws) {
    return *(this->set_str(ws));
}

xllType& xllType::operator=(const std::wstring& str) {
    return *(this->set_str(str));
}

xllType& xllType::operator=(const xlllist& list) {
    return *(this->set_list(list));
}

xllType& xllType::operator=(const xllmartix& matrix) {
    return *(this->set_matrix(matrix));
}

bool xllType::is_num() const {
    if (this->num != 0) return true;
    if (this->xltype == xltypeNum || this->xltype == xltypeInt) return true;
    if (this->optr && (this->optr->xltype == xltypeNum || this->optr->xltype == xltypeInt)) return true;
    return false;
}

bool xllType::is_str() const {
    if (!array.empty()) return true;
    if (this->xltype == xltypeStr) return true;
    if (this->optr && this->optr->xltype == xltypeStr) return true;
    return false;
}

bool xllType::is_array() const {
    if (!array.empty()) return true;
    if (this->xltype == xltypeMulti) return true;
    if (this->is_sref()) {
        int r = this->val.sref.ref.colLast - this->val.sref.ref.colFirst + 1;
        int c = this->val.sref.ref.rwLast - this->val.sref.ref.rwFirst + 1;
        if (r * c > 1) return true;
    }
    return false;
}

bool xllType::is_sref() const {
    return (this->xltype == xltypeSRef || this->xltype == xltypeRef) ? true : false;
}

xllType* xllType::serialize() {
    if (!this->is_array()) return this;
    if (this->array.empty()) return this;
    int size = this->array.size();
    std::vector<std::vector<std::wstring>> data;
    std::vector<std::wstring> row;
    int i = 0, n = (this->rows < 1 || size % this->rows > 0) ? 1 : this->rows;
    for (auto& x : this->array) {
        std::wstring t;
        if (x->is_num()) {
            // Number to string
            t = std::to_wstring(x->num);
            // Remove trailing zeros
            t.erase(t.find_last_not_of('0') + 1, std::wstring::npos);
            // If all digits after decimal point are zero, remove decimal point as well
            if (!t.empty() && t.back() == '.') t.pop_back();
        } else if (x->is_str()) {
            t = x->str;
        } else {
            t = L"";
        }
        row.emplace_back(std::move(t));
        i++;
        if (i >= n) {
            data.emplace_back(std::move(row));
            i = 0;
        }
    }
    std::wstring ret;
    if (!xllSerialize(data, ret)) return this;
    this->set_str(ret);
    return this;
}

xllType* xllType::deserialize() {
    if (!this->is_str() || str.empty()) return this;
    std::vector<std::vector<std::wstring>> ret;
    if (!xllDeserialize(this->str, ret)) return this;
    this->destory();
    this->rows = ret.size();
    this->cols = ret[0].size();
    for (auto& row : ret) {
        for (auto& col : row) {
            this->array.emplace_back(std::make_unique<xllType>(col));
        }
    }
    this->xltype = xltypeMulti;
    return this;
}

xllType* xllType::to_xloper12() {
    if (this->is_num()) {
        this->val.num = this->num;
    } else if (this->is_str()) {
        this->str.insert(0, 1, this->str.length());
        this->val.str = this->str.data();
    }
    return this;
}

xloper12* xllType::get_return() {
    xloper12* ret = new xloper12;
    ret->val = this->val;
    ret->xltype = this->xltype;
    if (this->is_array()) {
        int n = this->size();
        if (n <= 0) {
            ret->xltype = xltypeNil | xlbitDLLFree;
            return ret;
        }
        if (this->rows <= 0 || n % this->rows > 0) this->rows = 1;
        ret->val.array.rows = this->rows;
        ret->val.array.columns = int(n / this->rows);
        ret->val.array.lparray = new xloper12[n];
        for (int i = 0; i < n; i++) {
            ret->val.array.lparray[i] = *this->at(i)->get_return();
        }
        ret->xltype = xltypeMulti;
    } else if (this->is_str()) {
        ret->val.str = makeStr12(this->str.c_str());
        ret->xltype = xltypeStr;
    } else if (this->is_num()) {
        ret->val.num = this->num;
        ret->xltype = xltypeNum;
    }
    ret->xltype |= xlbitDLLFree;
    return ret;
}

xllType::Iter xllType::begin() {
    return Iter(this, 0);
}

xllType::Iter xllType::end() {
    return Iter(this, this->size());
}

bool xllType::Iter::operator!=(const xllType::Iter& r) const {
    return _i != r._i;
}

xllType* xllType::Iter::operator*() const {
    return _p->at(_i);
}

const xllType::Iter& xllType::Iter::operator++() {
    ++_i;
    return *this;
}

int xllType::size() const {
    return this->array.size();
}

xllType* xllType::at(int i) {
    int c = this->size();
    if (c == 0) return nullptr;
    this->pi = i < c ? i > 0 ? i : 0 : c - 1;
    return this->array.at(i).get();
}

xllType* xllType::at(int row, int col) {
    int _r = row < this->rows ? row < 1 ? 1 : row : this->rows;
    int _c = col < this->cols ? col < 1 ? 1 : col : this->cols;
    this->pi = ((_r - 1) * this->cols + _c) - 1;
    return this->array.at(this->pi).get();
}

xllType* xllType::operator[](int i) {
    return this->at(i);
}

xllType* xllType::push_back(const xllType& x) {
    if (!this->is_array()) {
        this->set_list({*this});
    }
    this->array.emplace_back(std::make_unique<xllType>(x));
    return this;
}

xllType* xllType::push_back(double num) {
    if (!this->is_array()) {
        this->set_list({*this});
    }
    this->array.emplace_back(std::make_unique<xllType>(num));
    return this;
}

xllType* xllType::push_back(const wchar_t* ws) {
    if (!this->is_array()) {
        this->set_list({*this});
    }
    this->array.emplace_back(std::make_unique<xllType>(ws));
    return this;
}

xllType* xllType::push_back(const std::wstring& ws) {
    if (!this->is_array()) {
        this->set_list({*this});
    }
    this->array.emplace_back(std::make_unique<xllType>(ws));
    return this;
}