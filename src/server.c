#define UNICODE
#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;

// {B672C9AD-3FE0-4712-A4BE-4525AD9A7F9F}
static const GUID SERVER_GUID  = 
{ 0xb672c9ad, 0x3fe0, 0x4712, { 0xa4, 0xbe, 0x45, 0x25, 0xad, 0x9a, 0x7f, 0x9f } };

static bool is_server_guid(GUID* test){
    return IsEqualGUID(test, &SERVER_GUID);
}

static void display_last_error(){
    DWORD last_error = GetLastError();
    u16* reason = NULL;

    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, last_error, 0, (LPWSTR)&reason, 1, NULL);
    MessageBoxW(NULL, reason, NULL, MB_OK);
    HeapFree(GetProcessHeap(), 0, reason);
}

static uint remove_last_part_from_path(u16* full_path, uint full_path_length){
    while (full_path_length >= 0){
        u16* current_ptr = full_path + full_path_length;
        u16 current_char = *current_ptr;
        if (current_char == L'\\'){
            return full_path_length + 1;
        } else {
            *current_ptr = 0;
            full_path_length -= 1;
        }
    }
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

#define MAX_UNICODE_PATH_LENGTH 32767

static u16 server_path[MAX_UNICODE_PATH_LENGTH + 1] = {0};
static uint server_path_length = 0;
static u16 server_directory[MAX_UNICODE_PATH_LENGTH + 1] = {0};
static uint server_directory_length = 0;

bool WINAPI entry_point(HINSTANCE hInst, DWORD reason, void* pReserved){
    if (reason == DLL_PROCESS_ATTACH){
        server_path_length = GetModuleFileNameW(hInst, (LPWSTR)&server_path, MAX_UNICODE_PATH_LENGTH);
        if (server_path_length == 0){
            display_last_error();
            return false;
        } else {
            MessageBoxW(NULL, server_path, L"server path", MB_OK);
            CopyMemory(&server_directory, &server_path, server_path_length * sizeof(u16));
            server_directory_length = remove_last_part_from_path(server_directory, server_path_length);
            MessageBoxW(NULL, server_directory, L"server directory", MB_OK);
        }
    }
    return true;
};



__declspec(dllexport) HRESULT __stdcall DllRegisterServer(void){
    MessageBoxW(NULL, L"Title", L"DllRegisterServer", MB_OK);
    return S_OK;
}

__declspec(dllexport) HRESULT __stdcall DllUnregisterServer(void){
    MessageBoxW(NULL, L"Title", L"DllUnregisterServer", MB_OK);
    return S_OK;
}
