echo OFF
cls
rm ./bin/extension.dll
rm ./obj/extension.obj

echo BUILDING EXTENSION
echo ------------------
gcc -shared -O2 -nostartfiles -Wall --std=c99 -c ./src/extension.c -o ./obj/extension.obj
ld -shared --kill-at --no-seh -e entry_point ./obj/extension.obj -o ./bin/extension.dll -L%LIBRARY_PATH% -lkernel32 -luuid -luser32 -lole32 -lNtosKrnl -lshell32 -lcomdlg32
strip ./bin/extension.dll
