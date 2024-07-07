project(Kiseki.ThumbnailHandler)

cmake_minimum_required(VERSION 3.23)

add_library(Kiseki.ThumbnailHandler SHARED
  src/thumb_win32.cc
  src/thumb_win32.def
  src/thumb_win32.rc
  src/thumb_win32_dll.cc
)

target_link_libraries(Kiseki.ThumbnailHandler dbghelp.lib Version.lib)