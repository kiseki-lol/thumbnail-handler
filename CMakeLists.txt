project(Kiseki.ThumbnailHandler)

cmake_minimum_required(VERSION 3.23)

add_library(Kiseki.ThumbnailHandler SHARED
  src/blendthumb_win32.cc
  src/blendthumb_win32.def
  src/blendthumb_win32.rc
  src/blendthumb_win32_dll.cc
)

target_link_libraries(Kiseki.ThumbnailHandler dbghelp.lib Version.lib)