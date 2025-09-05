cmake_minimum_required(VERSION 3.20)
project(EXCELXLL VERSION 0.1.0 LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED True)

include_directories(include)
link_directories(lib)
aux_source_directory(src SOURCES)

add_library(${PROJECT_NAME} SHARED
    functions.cpp dll.def
    ${SOURCES}
)

# 设置MSVC编译选项
if (MSVC)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /source-charset:utf-8 /execution-charset:utf-8")
    set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} /source-charset:utf-8 /execution-charset:utf-8")
endif()

if(CMAKE_SIZEOF_VOID_P EQUAL 4)
    # 32位支持WPS
    target_link_libraries(${PROJECT_NAME} XLCALL32) # 32位
    set_target_properties(${PROJECT_NAME} PROPERTIES PREFIX "" SUFFIX "_x86.xll")
elseif(CMAKE_SIZEOF_VOID_P EQUAL 8)
    # 64位支持Excel
    target_link_libraries(${PROJECT_NAME} x64/XLCALL32) # 64位
    set_target_properties(${PROJECT_NAME} PROPERTIES PREFIX "" SUFFIX "_x64.xll")
endif()

# 编译选项选择非Debug模式后开启静态编译，尽量减小编译文件的依赖和文件大小
if(NOT CMAKE_BUILD_TYPE STREQUAL "Debug")
    if (CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    
        target_compile_options(${PROJECT_NAME} PRIVATE
            -ffunction-sections
            -fdata-sections
        )
        set_target_properties(${PROJECT_NAME} PROPERTIES
            LINK_FLAGS "-flto -Wl,--gc-sections"
        )
        set(CMAKE_EXE_LINKER_FLAGS "-s -static")
        set(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -s -static")

    elseif (CMAKE_CXX_COMPILER_ID STREQUAL "MSVC")
        target_compile_options(${PROJECT_NAME} PRIVATE /MT)
    endif()
endif()