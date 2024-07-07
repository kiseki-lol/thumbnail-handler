

set(SRC_WIN32
  src/blendthumb_win32.cc
  src/blendthumb_win32.def
  src/blendthumb_win32.rc
  src/blendthumb_win32_dll.cc
)

add_library(Kiseki.ThumbnailHandler SHARED ${SRC} ${SRC_WIN32})

target_link_libraries(Kiseki.ThumbnailHandler dbghelp.lib Version.lib)
