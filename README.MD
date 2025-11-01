# Replorer

Маленькая утилита для Windows, которая перезапускает Проводник (explorer.exe):
1) Сначала пытается **аккуратно** завершить Explorer (команда `Exit Explorer`),
2) Если тот «упрямится» — **принудительно** завершает все процессы `explorer.exe`,
3) Сразу **запускает** новый экземпляр.

Полезно, когда пропадают/ломаются миниатюры иконок, глючит трэй или контекстные меню.

> **Внимание:** все окна Проводника будут закрыты. Сохраните работу в приложениях перед запуском.

## Сборка

Проект состоит из одного файла `main.cpp`. В CLion (CMake) добавьте минимальный `CMakeLists.txt` у себя локально:

```cmake
cmake_minimum_required(VERSION 3.20)
project(Replorer LANGUAGES CXX)
add_executable(Replorer main.cpp)
if (MSVC)
  target_compile_options(Replorer PRIVATE /W4 /utf-8)
else()
  target_compile_options(Replorer PRIVATE -Wall -Wextra -Wpedantic)
endif()
