static const GUID SERVER_GUID =
    { 0xb672c9ad, 0x3fe0, 0x4712, { 0xa4, 0xbe, 0x45, 0x25, 0xad, 0x9a, 0x7f, 0x9f } };

#define SERVER_GUID_TEXT  L"{B672C9AD-3FE0-4712-A4BE-4525AD9A7F9F}"
#define MAX_UNICODE_PATH_LENGTH 32767

static void display_error(DWORD error){
    u16* reason = NULL;

    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, error, 0, (LPWSTR)&reason, 1, NULL);
    MessageBoxW(NULL, reason, NULL, MB_OK);
    HeapFree(GetProcessHeap(), 0, reason);
}

static void display_last_error(){
    DWORD last_error = GetLastError();
    return display_error(last_error);
}

