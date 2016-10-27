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

#define PROGNAME L"Copy-Move Shell Extension In/Unstaller"
static const u16 extension_name[] = L"extension.dll\0";

static bool ask(USHORT* question){
    return IDYES == MessageBoxW(NULL, question, PROGNAME, MB_YESNO | MB_ICONQUESTION);
}

static void inform (USHORT* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONINFORMATION);
}

static void error(USHORT* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONERROR);
}

static void warn(USHORT* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONWARNING);
}

static bool register_com_server(u16* server_path, uint server_path_length){
    static u16 description[] = L"Windows Explorer context menu extension to copy, move and duplicate files and directories\0";

    HKEY clsidKey;
    LONG errorCode = RegCreateKeyExW(
         HKEY_CLASSES_ROOT,
         L"CLSID\\" SERVER_GUID_TEXT,
         0, NULL,
         REG_OPTION_NON_VOLATILE,
         KEY_ALL_ACCESS,
         NULL,
         &clsidKey,
         NULL
    );
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }
    errorCode = RegSetValueExW(clsidKey, NULL, 0, REG_SZ, (BYTE*)description, sizeof(description));
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }
    
    HKEY inprocKey;
    errorCode = RegCreateKeyExW(
              clsidKey,
              L"InprocServer32",
              0, NULL,
              REG_OPTION_NON_VOLATILE,
              KEY_ALL_ACCESS,
              NULL,
              &inprocKey,
              NULL              
    );
    
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }

    errorCode = RegSetValueExW(inprocKey, NULL, 0, REG_SZ, (BYTE*)server_path, (server_path_length + 1) * sizeof(u16));
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }

    errorCode = RegSetValueExW(inprocKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"both", sizeof(L"both") + 2);
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }
    

    RegCloseKey(inprocKey);
    RegCloseKey(clsidKey);
    return true;
}

static void register_shell_extension(){

};

static void uninstall_extension(){
    inform(L"uninstall");
}

static void install_extension(u16* server_path){
    inform(server_path);
    if (register_com_server(server_path, lstrlenW(server_path))){
        register_shell_extension();
    };

}

static bool is_extension_installed(){
    HRESULT result = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    IUnknown* pIUnknown = NULL;

    if (result != S_OK){
        error(L"Failed to initialize COM");
        return false;
    }
    result = CoCreateInstance(&SERVER_GUID, NULL, CLSCTX_INPROC_SERVER, &IID_IUnknown, (void**)&pIUnknown);
    if (result == S_OK){
        pIUnknown->lpVtbl->Release(pIUnknown);
        return true;
    }
    return false;
}

static u16 processFullPath[MAX_UNICODE_PATH_LENGTH + 1] = {0};
static uint processFullPathLength = 0;
void entry_point(){
    processFullPathLength = GetModuleFileNameW(NULL, processFullPath, MAX_UNICODE_PATH_LENGTH);
    if (processFullPathLength == 0){
        display_last_error();
        ExitProcess(1);
        return;
    }

    const u16* commandLine = GetCommandLineW();

    int nArgs = 0;
    u16** args = CommandLineToArgvW(commandLine, &nArgs);
    if (args == NULL){
        display_last_error();
        ExitProcess(1);
        return;
    }

    if (nArgs == 1){
        if (is_extension_installed()){
            if (ask(L"Shell extension is installed, would you like to UNINSTALL it?")){
                ShellExecuteW(NULL, L"runas", processFullPath, L"u", NULL, SW_SHOWDEFAULT);
            }
        } else {
            if (ask(L"Would you like to INSTALL the extension?")){
                ShellExecuteW(NULL, L"runas", processFullPath, L"i", NULL, SW_SHOWDEFAULT);
            }
        }

    } else {
        //one or more args were passed. take first arg and see what is it
        u16* arg = args[1];
        switch(*arg){
           case L'i':{
               uint length = remove_last_part_from_path(processFullPath, processFullPathLength);
               CopyMemory(processFullPath + length, extension_name, sizeof(extension_name));

               install_extension(processFullPath);
           } break;
               
           case L'u':{
               uninstall_extension();
           } break;
               
           default:{
               error(L"Invalid command line");
           }
        }
    }

    ExitProcess(0);
    return;
}
