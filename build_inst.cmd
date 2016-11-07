echo OFF
cls
rm ./obj/installer.obj
rm "./bin/installer.exe"

echo BUILDING INSTALLER
echo ------------------
gcc -O2 -nostartfiles -Wall --std=c99 -c ./src/installer.c -o ./obj/installer.obj
ld -s --subsystem windows --kill-at --no-seh -e entry_point ./obj/installer.obj -o "./bin/installer.exe"  -L%LIBRARY_PATH% -lkernel32 -luuid -luser32 -lole32 -lNtosKrnl -ladvapi32 -lshell32
