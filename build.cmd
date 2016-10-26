@echo off
cls
RD /S /Q obj
RD /S /Q bin
mkdir obj
mkdir bin

REM build server
gcc -shared -O2 -nostartfiles -Wall --std=c99 -c ./src/server.c -o ./obj/main.obj
ld -shared --kill-at --no-seh -e entry_point ./obj/main.obj -o ./bin/server.dll -L%LIBRARY_PATH% -lkernel32 -luuid -luser32 -lole32 -lNtosKrnl -ladvapi32
cp src/register.extension.cmd bin/