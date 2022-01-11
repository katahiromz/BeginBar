# CMakeLists.txt --- CMake project settings
#    ex) cmake -G "Visual Studio 9 2008" .
#    ex) cmake -DCMAKE_BUILD_TYPE=Release -G "MSYS Makefiles" .
##############################################################################

# CMake minimum version
cmake_minimum_required(VERSION 3.0)

# project name and languages
project(BeginBar CXX RC)

##############################################################################

# BeginBar.exe
add_executable(BeginBar WIN32
    BeginBar.cpp
    dialogs.cpp
    fakemenu.cpp
    settings.cpp
    tray.cpp
    BeginBar_res.rc)
target_link_libraries(BeginBar comctl32 shlwapi)
target_compile_definitions(BeginBar PRIVATE -DUNICODE -D_UNICODE)

##############################################################################
