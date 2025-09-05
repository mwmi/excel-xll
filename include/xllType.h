/**
 * @file xllType.h
 * @brief Excel XLL Plugin Core Data Type Encapsulation
 * @author mwm
 * @version 1.0
 * @date 2024-05-10
 * @copyright Copyright (c) 2024
 * 
 * This file defines the core data type xllType class of the ExcelXLL framework, which inherits from the xloper12 structure of the Excel C API,
 * providing type-safe, automatic memory management and convenient operation high-level interfaces for Excel XLL plugin development.
 * 
 * __Main Features__:
 * - __Type Safety__: Provides strong type interfaces to avoid direct manipulation of underlying xloper12 structures
 * - __Automatic Memory Management__: Based on RAII principles, automatically handles allocation and deallocation of strings and array memory
 * - __Rich Type Conversion__: Supports conversion between multiple Excel data types such as numbers, strings, arrays
 * - __Array Operation Support__: Provides iterators and index access, supports one-dimensional and two-dimensional array operations
 * - __Cell Reference Processing__: Automatically handles Excel cell references and circular reference detection
 * - __Serialization Support__: Supports serialization and deserialization of complex data structures
 * 
 * __Design Patterns__:
 * - __Inheritance Pattern__: Inherits xloper12, maintaining full compatibility with Excel C API
 * - __RAII Pattern__: Constructor allocates resources, destructor automatically releases resources
 * - __Iterator Pattern__: Provides standard begin/end iterator interfaces
 * - __Smart Pointer Management__: Uses unique_ptr to manage dynamically allocated memory
 * 
 * __Usage Scenarios__:
 * - UDF function parameter processing and return value construction
 * - Safe conversion between Excel data types
 * - Complex array and matrix operation processing
 * - RTD function data transfer and processing
 * 
 * @warning This class object will be automatically released by Excel after UDF function returns, please ensure proper use of get_return() method
 * @note Supports move semantics for efficient transfer of large array data
 * 
 * @see xloper12 Excel C API basic data structure
 * @see UDF Excel user-defined function macro
 * @see RTD Excel real-time data function
 */
#pragma once
#include "XLCALL.H"
#include <vector>
#include <string>
#include <memory>

class xllType;

/// @brief xloper12 smart pointer type, used for automatic memory management
using xlptr = std::unique_ptr<xloper12>;

/// @brief xllType smart pointer type, used for managing xllType objects
using xllptr = std::unique_ptr<xllType>;

/// @brief xllType smart pointer list, used for managing xllType object arrays
using xllptrlist = std::vector<xllptr>;

/// @brief xllType object list, used for constructing one-dimensional arrays
using xlllist = std::vector<xllType>;

/// @brief xllType two-dimensional matrix, used for constructing two-dimensional arrays
using xllmartix = std::vector<xlllist>;

/**
 * @brief Excel XLL Plugin Core Data Type Encapsulation Class
 * 
 * The xllType class is the core component of the ExcelXLL framework, inheriting from the xloper12 structure of the Excel C API,
 * providing advanced, type-safe data encapsulation interfaces for Excel XLL plugin development.
 * 
 * __Core Features__:
 * 
 * __Type Safety__:
 * - Provides strong type interfaces, avoiding direct manipulation of complex xloper12 structures
 * - Automatic type detection and conversion, supporting methods like is_num(), is_str(), is_array()
 * - Prevents type errors and data corruption
 * 
 * __Memory Management__:
 * - Based on RAII principles, automatically allocates on construction, automatically releases on destruction
 * - Uses std::unique_ptr to manage dynamic memory, preventing memory leaks
 * - Supports move semantics for efficient large data transfer
 * 
 * __Data Type Support__:
 * - __Numeric Types__: Supports double type numeric storage and conversion
 * - __String Types__: Supports Unicode strings, compatible with std::wstring and C-style strings
 * - __Array Types__: Supports one-dimensional and two-dimensional arrays, provides iterator and index access
 * - __Cell References__: Supports Excel cell reference types, automatically parses reference content
 * - __Error Types__: Supports Excel error types, provides error information transfer
 * 
 * __Advanced Features__:
 * - __Circular Reference Detection__: Prevents circular reference issues in UDF functions
 * - __Data Serialization__: Supports serialization and deserialization of complex data structures
 * - __Iterator Support__: Provides STL-style iterators, supports range-based for loops
 * - __Multi-dimensional Array Access__: Supports at(row, col) method to access 2D array elements
 * 
 * __Usage Examples__:
 * 
 * ```cpp
 * // 1. Basic numeric operations
 * xllType num(3.14);
 * xllType result = num.get_num() * 2;
 * 
 * // 2. String processing
 * xllType str(L"Hello World");
 * std::wstring content = str.get_str();
 * 
 * // 3. Array operations
 * xllType arr;
 * arr = {xllType(1.0), xllType(2.0), xllType(3.0)};
 * for (auto item : arr) {
 *     double val = item->get_num();
 * }
 * 
 * // 4. Using in UDF functions
 * UDF(MyFunction, L"Test function", Param input) {
 *     xllType result;
 *     xllType input_data = input;
 *     if (input_data.is_num()) {
 *         result = input_data.get_num() * 2;
 *     } else {
 *         result = L"Type error";
 *     }
 *     return result.get_return();
 * }
 * ```
 * 
 * @note Inherits from xloper12, maintaining full compatibility with Excel C API
 * @warning When using in UDF functions, must return results through get_return() method
 * @warning Avoid modifying the same xllType object simultaneously in multi-threaded environments
 * 
 * @see xloper12 Excel C API basic data structure
 * @see UDF User-defined function macro
 * @see xlltools.h Utility function library
 */
class xllType : public xloper12 {
private:
    /// @brief Current array access index, used to record the last accessed array position
    int pi = 0;
    
    /// @brief Array row count, used for row dimension management of 2D arrays
    int rows = 0;
    
    /// @brief Array column count, used for column dimension management of 2D arrays
    int cols = 0;
    
    /// @brief Error code, records error information from the last operation
    int error_code = 0;
    
    /// @brief Numeric storage, used to cache numeric values parsed from xloper12
    double num = 0;
    
    /// @brief Excel data smart pointer, used to manage temporary data obtained from Excel
    xlptr optr = nullptr;
    
    /// @brief String storage, used to cache string data parsed from xloper12
    std::wstring str;
    
    /// @brief Array element storage, using smart pointers to manage array element memory
    xllptrlist array;
    
    /**
     * @brief Initialize object to empty state
     * @return xllType* Returns current object pointer to support method chaining
     * 
     * Sets the object's xltype to xltypeNil and clears all data members.
     * This method is called in the constructor to ensure consistency of the object's initial state.
     */
    xllType* init();
    
    /**
     * @brief Initialize object from xloper12 structure
     * @param x Source xloper12 structure reference
     * @return xllType* Returns current object pointer to support method chaining
     * 
     * Copies the xltype and val members of the xloper12 structure, then calls the load() method
     * to parse specific data content. This method is used to handle data passed from Excel.
     */
    xllType* init(const xloper12& x);
    
    /**
     * @brief Destroy object and release all resources
     * @return xllType* Returns current object pointer to support method chaining
     * 
     * Clears string, array and smart pointer members, releasing all dynamically allocated memory.
     * Called in destructor and reassignment operations to ensure memory safety.
     */
    xllType* destory();
    
    /**
     * @brief Load and parse xloper12 data content
     * @return xllType* Returns current object pointer to support method chaining
     * 
     * Parses and loads corresponding data content based on the type of xloper12.
     * For cell reference types, calls Excel API to get actual data.
     * For array types, checks for circular references and parses array elements.
     */
    xllType* load();
    
    /**
     * @brief Copy all data from another xllType object
     * @param other Source xllType object reference
     * @return xllType* Returns current object pointer to support method chaining
     * 
     * Performs deep copy, including basic data members and array elements.
     * Each element in the array will create a new xllType object copy.
     * This method is used in the copy constructor.
     */
    xllType* copy(const xllType& x);
    
    /**
     * @brief Load cell reference data
     * @param type Target type to convert to (xltypeNum, xltypeStr, xltypeMulti, etc.)
     * @return bool Returns true on success, false on failure
     * 
     * Uses Excel12(xlCoerce) API to convert cell references to specified type data.
     * For array types, parses all array elements and stores them in the array member.
     * Records error code to error_code member on failure.
     */
    bool load_ref(DWORD type);
    
    /**
     * @brief Check for circular references
     * @return bool Returns true if no circular reference detected, false otherwise
     * 
     * Checks if there's overlap with the reference area by getting current calling cell location.
     * If circular reference is detected, it may cause Excel calculation errors or infinite loops.
     * This method is used to prevent infinite recursion issues in UDF functions.
     */
    bool check_ref();
public:
    
    /// @name Construction and Destruction Functions
    
    /**
     * @brief Default constructor
     * 
     * Creates an empty xllType object, initialized as xltypeNil type.
     * All data members are set to default values and can be initialized later through assignment operators.
     * 
     * __Usage Examples__:
     * ```cpp
     * xllType obj;           // Create empty object
     * obj = 3.14;           // Assign as number
     * obj = L"Hello";        // Assign as string
     * ```
     */
    xllType();
    
    /**
     * @brief Destructor
     * 
     * Automatically releases all dynamically allocated resources, including string memory and array elements.
     * Based on RAII principles, ensures automatic resource cleanup when object lifetime ends,
     * Prevent memory leaks and resource leak problems.
     * 
     * @note Destructor is called automatically, no need to call manually
     */
    ~xllType();
    
    /**
     * @brief Copy constructor
     * @param other Source object to copy
     * 
     * Performs deep copy, creating a completely independent copy from source object.
     * For array types, recursively copies all elements to ensure data independence.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType original(3.14);
     * xllType copy(original);   // Copy construction
     * copy = 2.71;             // Modifying copy doesn't affect original object
     * ```
     */
    xllType(const xllType& other);
    
    /**
     * @brief Move constructor
     * @param other Source object to move (rvalue reference)
     * 
     * Efficiently transfers ownership of source object's resources, avoiding unnecessary copy operations.
     * Suitable for efficient transfer of large arrays or complex data structures.
     * After move, source object will be in a safe empty state.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType createArray() {
     *     xllType arr;
     *     arr = {xllType(1), xllType(2), xllType(3)};
     *     return arr;  // Automatically triggers move construction
     * }
     * xllType result = createArray();
     * ```
     */
    xllType(xllType&& other) noexcept;
    
    /**
     * @brief Construct from xloper12 pointer
     * @param px Pointer to xloper12 structure
     * 
     * Creates xllType object from LPXLOPER12 pointer passed from Excel.
     * Mainly used for UDF function parameter initialization, automatically parses data from Excel.
     * Supports all Excel data types, including numbers, strings, arrays and cell references.
     * 
     * __Usage Example__:
     * ```cpp
     * UDF(MyFunction, L"Test", Param input) {
     *     xllType data(input);  // Create from Excel parameter
     *     return data.get_return();
     * }
     * ```
     * 
     * @warning Ensure the passed pointer is non-null and valid
     */
    xllType(const xloper12* px);
    
    /**
     * @brief Construct from xloper12 structure
     * @param x xloper12 structure reference
     * 
     * Creates xllType object from Excel's xloper12 structure.
     * Similar to pointer version, but can accept structure reference parameters.
     * Automatically recognizes and parses different types of Excel data.
     * 
     * __Usage Example__:
     * ```cpp
     * xloper12 excelData;
     * // ... initialize excelData
     * xllType data(excelData);
     * ```
     */
    xllType(const xloper12& x);
    
    /**
     * @brief Construct from numeric value
     * @param num Double precision floating point number
     * 
     * Creates an xllType object containing the specified numeric value.
     * Object type will be set to xltypeNum, supports implicit conversion from all C++ numeric types.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType pi(3.14159);
     * xllType count(42);
     * xllType result = pi.get_num() * count.get_num();
     * ```
     */
    xllType(double num);
    
    /**
     * @brief Construct from C-style string
     * @param str Null-terminated ANSI string
     * 
     * Creates an xllType object containing the specified string.
     * ANSI string will be automatically converted to Unicode string for storage.
     * Object type will be set to xltypeStr.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType text("Hello World");
     * std::wstring content = text.get_str();
     * ```
     * 
     * @note Unicode version is recommended for better internationalization support
     */
    xllType(const char* str);
    
    /**
     * @brief Construct from Unicode string
     * @param ws Null-terminated Unicode string
     * 
     * Creates an xllType object containing the specified Unicode string.
     * Supports all Unicode characters, including Chinese, special symbols, etc.
     * Object type will be set to xltypeStr.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType chinese(L"Hello World");
     * xllType symbol(L"παβγ");
     * ```
     */
    xllType(const wchar_t* ws);
    
    /**
     * @brief Construct from std::wstring
     * @param str std::wstring string reference
     * 
     * Creates an xllType object containing the specified std::wstring string.
     * Provides seamless integration with STL string classes, supports all std::wstring features.
     * Object type will be set to xltypeStr.
     * 
     * __Usage Example__:
     * ```cpp
     * std::wstring msg = L"Dynamically generated message";
     * xllType result(msg);
     * ```
     */
    xllType(const std::wstring& str);

    xllType(const xllmartix &m);
    xllType(const xlllist &l);
    
    
    /// @name Assignment Operators
    
    /**
     * @brief Copy assignment operator
     * @param other Source object to copy
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Performs deep copy, first clears all data in current object, then copies all content from source object.
     * Supports self-assignment detection, for array types recursively copies all elements.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType a(3.14), b(2.71);
     * a = b;                    // Copy assignment
     * xllType c = a = b;        // Chained assignment
     * ```
     */
    xllType& operator=(const xllType& other);
    
    /**
     * @brief Assign from xloper12
     * @param x Source xloper12 structure reference
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Clears current object and reinitializes from xloper12 structure.
     * Automatically recognizes and parses different types of Excel data, supports automatic parsing of cell references.
     * 
     * __Usage Example__:
     * ```cpp
     * xloper12 excelData;
     * // ... initialize excelData
     * xllType obj;
     * obj = excelData;          // Assign from Excel data
     * ```
     */
    xllType& operator=(const xloper12& x);
    
    /**
     * @brief Assign from numeric value
     * @param num Double precision floating point number
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Sets object to numeric type and stores the specified value.
     * Clears all previous data and reinitializes as numeric type.
     * Supports implicit conversion from all C++ numeric types.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType obj;
     * obj = 3.14;              // Assign as floating point
     * obj = 42;                // Assign as integer (automatic conversion)
     * ```
     */
    xllType& operator=(double num);
    
    /**
     * @brief Assign from Unicode string
     * @param ws Null-terminated Unicode string
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Sets object to string type and stores the specified Unicode string.
     * Clears all previous data and reinitializes as string type.
     * Supports all Unicode characters, including Chinese, special symbols, etc.
     * 
     * __Usage Example__:
     * ```cpp
     * xllType obj;
     * obj = L"Hello World";        // Unicode text
     * obj = L"παβγ";           // Greek letters
     * ```
     */
    xllType& operator=(const wchar_t* ws);
    
    /**
     * @brief Assign from std::wstring
     * @param str std::wstring string reference
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Sets object to string type and stores the specified std::wstring string.
     * Provides seamless integration with STL string classes, supports all std::wstring features.
     * Clears all previous data and reinitializes as string type.
     * 
     * __Usage Example__:
     * ```cpp
     * std::wstring msg = L"Dynamic generation";
     * xllType obj;
     * obj = msg;               // Assign from STL string
     * ```
     */
    xllType& operator=(const std::wstring& str);
    
    /**
     * @brief Assign from one-dimensional array
     * @param list One-dimensional list of xllType objects
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Sets object to array type and stores the specified one-dimensional array data.
     * Array will be set as 1 row N columns matrix form, each element will be copied.
     * Clears all previous data and reinitializes as array type.
     * 
     * __Usage Example__:
     * ```cpp
     * xlllist data = {xllType(1.0), xllType(2.0), xllType(L"Text")};
     * xllType obj;
     * obj = data;              // Assign as one-dimensional array
     * ```
     */
    xllType& operator=(const xlllist& list);
    
    /**
     * @brief Assign from two-dimensional matrix
     * @param matrix Two-dimensional matrix of xllType objects
     * @return xllType& Returns reference to current object to support chained assignment
     * 
     * Sets object to array type and stores the specified two-dimensional matrix data.
     * Matrix row and column counts will be correctly recorded, supports at(row, col) access method.
     * Clears all previous data and reinitializes as array type.
     * 
     * __Usage Example__:
     * ```cpp
     * xllmartix matrix = {
     *     {xllType(1), xllType(2)},
     *     {xllType(3), xllType(4)}
     * };
     * xllType obj;
     * obj = matrix;            // Assign as two-dimensional matrix
     * ```
     */
    xllType& operator=(const xllmartix& matrix);
    
    /// @name Setter Functions
    
    /// @brief Set error code
    /// @param err Excel error code (such as xlerrRef, xlerrValue, etc.)
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_err(int err);
    
    /// @brief Set numeric value
    /// @param num Double precision floating point number
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_num(double num);
    
    /// @brief Set Unicode string
    /// @param ws Null-terminated Unicode string
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_str(const wchar_t* ws);
    
    /// @brief Set std::wstring string
    /// @param ws std::wstring string reference
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_str(const std::wstring& ws);
    
    /// @brief Set one-dimensional array
    /// @param l One-dimensional list of xllType objects
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_list(const xlllist& l);
    
    /// @brief Set two-dimensional matrix
    /// @param m Two-dimensional matrix of xllType objects
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* set_matrix(const xllmartix& m);
    
    
    /// @name Getter Functions
    
    /// @brief Get numeric value
    /// @return double Returns stored numeric value
    double get_num() const;
    
    /// @brief Get string
    /// @return std::wstring Returns string content
    std::wstring get_str() const;
    
    /// @brief Get C-style string pointer
    /// @return const wchar_t* Returns pointer to string
    const wchar_t* get_c_str() const;
    
    /// @brief Get last error code
    /// @return int Returns error code, 0 for normal
    int get_last_err() const;
    
    
    /// @name Type Checking Functions
    
    /// @brief Check if it's numeric type
    /// @return bool Returns true if numeric type, false otherwise
    bool is_num() const;
    
    /// @brief Check if it's string type
    /// @return bool Returns true if string type, false otherwise
    bool is_str() const;
    
    /// @brief Check if it's array type
    /// @return bool Returns true if array type, false otherwise
    bool is_array() const;
    
    /// @brief Check if it's cell reference type
    /// @return bool Returns true if cell reference, false otherwise
    bool is_sref() const;
    
    
    /// @name Serialization Functions
    
    /// @brief Serialize array to string
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* serialize();
    
    /// @brief Deserialize from string to array
    /// @return xllType* Returns current object pointer to support chained calls
    xllType* deserialize();
    
    /// @brief Convert to variable suitable for passing xloper12 type
    /// @return xllType* Returns current object pointer to support chained calls
    /// @warning This operation may cause string usage exceptions, usually used at the end, not recommended for regular operations after use
    xllType* to_xloper12();
    
    /// @brief Get Excel return value pointer
    /// @return xloper12* Returns xloper12 pointer recognizable by Excel
    /// @warning The returned pointer is released by Excel, do not manually delete
    xloper12* get_return();
    
    /**
     * @brief STL-style iterator class
     * 
     * Provides STL-compliant iterator interface, supports range-based for loops.
     * Used to traverse each element in xllType array, returns xllType pointer.
     */
    class Iter {
    private:
        /// @brief Current iteration position index
        int _i;
        /// @brief Pointer to parent xllType object
        xllType* _p;
    public:
        /// @brief Constructor
        /// @param p Pointer to xllType object
        /// @param i Initial iteration position
        Iter(xllType* p, int i) : _i(i), _p(p) {}
        
        /// @brief Inequality comparison operator
        bool operator!=(const Iter& r) const;
        
        /// @brief Dereference operator
        /// @return xllType* Returns xllType pointer at current position
        xllType* operator*() const;
        
        /// @brief Prefix increment operator
        const Iter& operator++();
    };
    
    
    /// @name Iterator and Array Access Functions
    
    /// @brief Get iterator starting position
    Iter begin();
    
    /// @brief Get iterator ending position
    Iter end();
    
    /// @brief Get array element count
    int size() const;
    
    /// @brief Access array element by index
    /// @param i Index position
    /// @return xllType* Returns element pointer at specified position
    xllType* at(int i);
    
    /// @brief Access 2D array element by row and column index
    /// @param row Row index (1-based indexing)
    /// @param col Column index (1-based indexing)
    /// @return xllType* Returns element pointer at specified position
    xllType* at(int row, int col);
    
    /// @brief Array subscript access operator
    /// @param i Index position
    /// @return xllType* Returns element pointer at specified position
    xllType* operator[](int i);
    
    /// @brief Add element
    /// @param x Element to add
    /// @return xllType* Returns current object pointer
    xllType* push_back(const xllType& x);
    xllType* push_back(double num);
    xllType* push_back(const wchar_t* ws);
    xllType* push_back(const std::wstring& ws);
    
};