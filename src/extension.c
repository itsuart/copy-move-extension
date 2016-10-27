#define UNICODE
#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;

/* GLOBALS */
#include "common.c"

static bool is_server_guid(GUID* test){
    return IsEqualGUID(test, &SERVER_GUID);
}

static HANDLE hLogFile = INVALID_HANDLE_VALUE;
static void open_log_file(){
    if (hLogFile != INVALID_HANDLE_VALUE){
        return;
    }
    hLogFile = CreateFileW( L"C:\\tmp\\server_log.txt",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    
}

static void close_log_file(){
    if (hLogFile != INVALID_HANDLE_VALUE){
        CloseHandle(hLogFile);
        hLogFile = INVALID_HANDLE_VALUE;
    }
}



bool WINAPI entry_point(HINSTANCE hInst, DWORD reason, void* pReserved){
    return true;
};

__declspec(dllexport) HRESULT __stdcall DllRegisterServer(void){
    return S_OK;
}

