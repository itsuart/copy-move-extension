@echo off
cls
RD /S /Q obj
RD /S /Q bin
mkdir obj
mkdir bin

REM build extension
gcc -shared -O2 -nostartfiles -Wall --std=c99 -c ./src/extension.c -o ./obj/extension.obj
ld -shared --kill-at --no-seh -e entry_point ./obj/extension.obj -o ./bin/extension.dll -L%LIBRARY_PATH% -lkernel32 -luuid -luser32 -lole32 -lNtosKrnl -ladvapi32

REM build installer
gcc -O2 -nostartfiles -Wall --std=c99 -c ./src/installer.c -o ./obj/installer.obj
ld -s --subsystem windows --kill-at --no-seh -e entry_point ./obj/installer.obj -o "./bin/in staller.exe"  -L%LIBRARY_PATH% -lkernel32 -luuid -luser32 -lole32 -lNtosKrnl -ladvapi32 -lshell32