#define _WIN32_WINNT _WIN32_WINNT_WIN7
#define UNICODE

#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>
#include <Shobjidl.h>
#include <Shlobj.h>

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;
typedef unsigned char u8;

//#define ODS(x) OutputDebugStringW(x)
#define ODS(x) (void)0

/* GLOBALS */
#include "common.c"

/* FORWARD DECLARATION EVERYTHING */
struct tag_IMyObj;

ULONG STDMETHODCALLTYPE myObj_AddRef(struct tag_IMyObj* pMyObj);
ULONG STDMETHODCALLTYPE myObj_Release(struct tag_IMyObj* pMyObj);
HRESULT STDMETHODCALLTYPE myObj_QueryInterface(struct tag_IMyObj* pMyObj, REFIID requestedIID, void **ppv);

/* ALL THE STRUCTS (because it's too much pain to have them in separate places) */
typedef struct tag_MyIContextMenuImpl {
    IContextMenuVtbl* lpVtbl;
    struct tag_IMyObj* pBase;
} MyIContextMenuImpl;

typedef struct tag_MyIShellExtInitImpl {
    IShellExtInitVtbl* lpVtbl;
    struct tag_IMyObj* pBase;
} MyIShellExtInitImpl;

typedef struct tag_IMyObj{
    IUnknownVtbl* lpVtbl;
    MyIContextMenuImpl contextMenuImpl;
    MyIShellExtInitImpl shellExtInitImpl;
    uint nItems;
    u16* itemNames;
    long int refsCount;
} IMyObj;

static void cleanup_names_storage(IMyObj* pBase){
    if (pBase->itemNames != NULL){
        GlobalFree(pBase->itemNames);
    }
    pBase->itemNames = NULL;
    pBase->nItems = 0;
}



/******************************* IContextMenu ***************************/
static const u16* COPY_VERB = SERVER_GUID_TEXT L".COPY";
static const u16* MOVE_VERB = SERVER_GUID_TEXT L".MOVE";

#define COPY_TO_MENU_OFFSET 1
#define MOVE_TO_MENU_OFFSET 2

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_GetCommandString(MyIContextMenuImpl* pImpl, 
                                                              UINT_PTR idCmd, UINT uFlags, UINT* pwReserved,LPSTR pszName,UINT cchMax){
    if (uFlags == GCS_VERBW){
        switch (idCmd){
            case COPY_TO_MENU_OFFSET:{
                lstrcpyW((u16*)pszName, COPY_VERB);
                ODS(L"GetCommandString SUCCESS (copy verb)");
                return S_OK;
            } break;

            case MOVE_TO_MENU_OFFSET:{
                lstrcpyW((u16*)pszName, MOVE_VERB);
                ODS(L"GetCommandString SUCCESS (move verb)");
                return S_OK;
            } break;

        }
    } else if (uFlags == GCS_VALIDATEW){
        if (idCmd == COPY_TO_MENU_OFFSET || idCmd == MOVE_TO_MENU_OFFSET){
            ODS(L"GetCommandString SUCCESS (validate cmd)");
            return S_OK;
        } else {
            ODS(L"GetCommandString UNKNOWN CMD VALIDATION");
            return S_FALSE;
        }
    }
    ODS(L"GetCommandString FAILED");
    return E_INVALIDARG;
}

//returns number of digits in the number
static uint u64toW(uint number, u16 stringBuffer[20]){
    const static u16 DIGITS[10] = L"0123456789";
    if (number == 0){ //special case
        stringBuffer[0] = L'0';
        return 1;
    }

    //biggest 64 bit unsigned number is 18,446,744,073,709,551,615 that is 20 digits, so
    uint divisor = 10;
    int power = 0;
    u8 reminders[20] = {0};

    for (; power < 20; power += 1){
        char reminder = number % divisor;

        reminders[power] = reminder;
        number = number / divisor;

        if (number == 0){
            break;
        }
    }

    //put 'em digits in correct order
    for (int i = power; i >= 0; i -= 1){
        stringBuffer[power - i] = DIGITS[reminders[i]];
    }

    return power;
}

static void mk_label (u16* prefix, uint number, u16* buffer){
    uint index = 0;
    while (prefix[index] != 0){
        buffer[index] = prefix[index];
        index += 1;
    }
    uint nDigits = u64toW(number, buffer + index);
    index += nDigits;

    lstrcpyW(buffer + index + 1, number > 1 ? L" items to..." : L" item to...");
}


//I really hate person that causes me "undefined reference to `___chkstk_ms'" errors. And extra calls to clear the memory.
static u16 targetDir[MAX_UNICODE_PATH_LENGTH + 2]; //extra 1 for trailing zero
HRESULT STDMETHODCALLTYPE myIContextMenuImpl_InvokeCommand(MyIContextMenuImpl* pImpl, LPCMINVOKECOMMANDINFO pCommandInfo){
    IMyObj* pBase = pImpl->pBase;
    
    ODS(L"FOO");

    if (pBase->nItems == 0){
        return S_OK; //instantly copy or move nothing!
    }
    ODS(L"BAR");

    void* pVerb = pCommandInfo->lpVerb;

    if (HIWORD(pVerb)){
        ODS(L"VERB");
        return E_INVALIDARG;
    }

    ODS(L"BAZ");

    u16 title[(sizeof(L"COPY ")/sizeof(L' ')) + 20 + (sizeof(L" items to...")/sizeof(L' ')) + 1];
    SecureZeroMemory(&title[0], sizeof(title));    
    uint fileOp;

    if ((void*)pVerb == (void*)MAKEINTRESOURCE(COPY_TO_MENU_OFFSET)){
        mk_label(L"COPY ", pBase->nItems, &title[0]);
        fileOp = FO_COPY;
    } else if ((void*)pVerb == (void*)MAKEINTRESOURCE(MOVE_TO_MENU_OFFSET)){
        mk_label(L"MOVE ", pBase->nItems, &title[0]);
        fileOp = FO_MOVE;
    } else {
        ODS(L"Neither COPY nor MOVE offset");
        cleanup_names_storage(pBase);
        return E_FAIL;
    }

    ODS(L"BAZ+");
    //I want my monads in C!
    HRESULT hr = S_OK; 
    // Create a new common open file dialog.
    IFileOpenDialog *pfd = NULL;
    hr = CoCreateInstance(&CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileOpenDialog, (void**)&pfd);
    if (SUCCEEDED(hr)){
        ODS(L"BAZ++");
        // Set the dialog as a folder picker.
        DWORD dwOptions;
        hr = pfd->lpVtbl->GetOptions(pfd, &dwOptions);
        if (SUCCEEDED(hr)){
            hr = pfd->lpVtbl->SetOptions(pfd, dwOptions | FOS_PICKFOLDERS);

            // Set the title of the dialog.
            if (SUCCEEDED(hr)){
                hr = pfd->lpVtbl->SetTitle(pfd, &title[0]);

                // Show the open file dialog.
                if (SUCCEEDED(hr)){
                    hr = pfd->lpVtbl->Show(pfd, pCommandInfo->hwnd);
                    if (SUCCEEDED(hr)){
                        // Get the selection from the user.
                        IShellItem *psiResult = NULL;
                        hr = pfd->lpVtbl->GetResult(pfd, &psiResult);
                        if (SUCCEEDED(hr)){
                            u16* pszPath = NULL;
                            hr = psiResult->lpVtbl->GetDisplayName(psiResult, SIGDN_FILESYSPATH, &pszPath);
                            if (SUCCEEDED(hr)){
                                SecureZeroMemory(targetDir, sizeof(targetDir));
                                lstrcpyW(&targetDir[0], pszPath);
                                CoTaskMemFree(pszPath);

                                SHFILEOPSTRUCT shFileOp = {
                                    .hwnd = pCommandInfo->hwnd,
                                    .wFunc = fileOp,
                                    .pFrom = pBase->itemNames,
                                    .pTo = &targetDir[0],
                                    .fFlags = FOF_ALLOWUNDO,
                                    .fAnyOperationsAborted = false,
                                    .hNameMappings = NULL,
                                    .lpszProgressTitle = L"ERROR!"
                                };

                                SHFileOperation(&shFileOp);                        
                            }
                            psiResult->lpVtbl->Release(psiResult);
                        }
                    }
                }
            }
        }
        pfd->lpVtbl->Release(pfd);        
    }

    cleanup_names_storage(pBase);
    return S_OK;
}


HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryContextMenu(MyIContextMenuImpl* pImpl, 
                                                           HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags){
    IMyObj* pBase = pImpl->pBase;
    
    if (CMF_DEFAULTONLY == (CMF_DEFAULTONLY & uFlags)
        || CMF_VERBSONLY == (CMF_VERBSONLY & uFlags)
        || CMF_NOVERBS == (CMF_NOVERBS & uFlags)){

        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //in highly unlikely case where is no room left in the menu:
    if (idCmdFirst + 2 > idCmdLast){
        //we don't add any menu enries of ours
        cleanup_names_storage(pBase);
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //insert our menu items: separator, copy and move in that order.
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + COPY_TO_MENU_OFFSET, L"Copy To...");
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + MOVE_TO_MENU_OFFSET, L"Move To...");

    ODS(L"QueryContextMenu");

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, MOVE_TO_MENU_OFFSET + 1); //shouldn't this be +idCmdFirst?
}


ULONG STDMETHODCALLTYPE myIContextMenuImpl_AddRef(MyIContextMenuImpl* pImpl){
    return myObj_AddRef(pImpl->pBase);
}

ULONG STDMETHODCALLTYPE myIContextMenuImpl_Release(MyIContextMenuImpl* pImpl){
    return myObj_Release(pImpl->pBase);
}

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryInterface(MyIContextMenuImpl* pImpl, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pImpl;
        myObj_AddRef(pImpl->pBase);
        return S_OK;
    }

    return myObj_QueryInterface(pImpl->pBase, requestedIID, ppv);

}

static IContextMenuVtbl IMyIContextMenuVtbl = {
    .QueryInterface = &myIContextMenuImpl_QueryInterface,
    .AddRef = &myIContextMenuImpl_AddRef,
    .Release = &myIContextMenuImpl_Release,
    .GetCommandString = &myIContextMenuImpl_GetCommandString,
    .QueryContextMenu = &myIContextMenuImpl_QueryContextMenu,
    .InvokeCommand = &myIContextMenuImpl_InvokeCommand
};



/************************ IShellExtInit **********************************/
HRESULT STDMETHODCALLTYPE myIShellExtInit_AddRef(MyIShellExtInitImpl* pImpl){
    return myObj_AddRef(pImpl->pBase);
}

ULONG STDMETHODCALLTYPE myIShellExtInit_Release(MyIShellExtInitImpl* pImpl){
    return myObj_Release(pImpl->pBase);
}

HRESULT STDMETHODCALLTYPE myIShellExtInit_QueryInterface(MyIShellExtInitImpl* pImpl, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pImpl;
        myObj_AddRef(pImpl->pBase);
        return S_OK;
    }

    return myObj_QueryInterface(pImpl->pBase, requestedIID, ppv);

}

HRESULT STDMETHODCALLTYPE myIShellExtInit_Initialize(MyIShellExtInitImpl* pImpl, LPCITEMIDLIST pIDFolder, IDataObject* pDataObj, HKEY hRegKey){
    FORMATETC   fe = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM   medium;

    if (pImpl->pBase->itemNames != NULL){
        GlobalFree(pImpl->pBase->itemNames);
        pImpl->pBase->itemNames = NULL;
        pImpl->pBase->nItems = 0;
    }

    if (SUCCEEDED(pDataObj->lpVtbl->GetData(pDataObj, &fe, &medium))){
        // Get the count of files dropped.
        uint uCount = DragQueryFile((HDROP)medium.hGlobal, (UINT)-1, NULL, 0);
        if (uCount > 0){
            u16* allNamesBuffer = GlobalAlloc(GPTR, (uCount * (MAX_UNICODE_PATH_LENGTH + 1) + 1) * sizeof(u16));
            if (allNamesBuffer == NULL){
                ReleaseStgMedium(&medium);
                return E_OUTOFMEMORY;
            }

            //copy file items to the buffer
            u16* currentName = allNamesBuffer;
            for (uint i = 0; i < uCount; i += 1){
                uint nCharsCopied = DragQueryFileW((HDROP)medium.hGlobal, i, currentName, MAX_UNICODE_PATH_LENGTH);
                if (nCharsCopied){
                    currentName += (nCharsCopied + 1);
                } else {/*copying fucked up */}
            }
            pImpl->pBase->nItems = uCount;
            pImpl->pBase->itemNames = allNamesBuffer;
        }

        ReleaseStgMedium(&medium);
    }

    ODS(L"Initialize success");

    return S_OK;
}

static IShellExtInitVtbl IMyIShellExtVtbl = {
    .AddRef = &myIShellExtInit_AddRef,
    .Release = &myIShellExtInit_Release,
    .QueryInterface = &myIShellExtInit_QueryInterface,
    .Initialize = &myIShellExtInit_Initialize
};



/******************* IMyObj  *******************************/
static long int nObjectsAndRefs = 0;

ULONG STDMETHODCALLTYPE myObj_AddRef(IMyObj* pMyObj){
    InterlockedIncrement(&nObjectsAndRefs);

    InterlockedIncrement(&(pMyObj->refsCount));
    return pMyObj->refsCount;
}

ULONG STDMETHODCALLTYPE myObj_Release(IMyObj* pMyObj){
    InterlockedDecrement(&nObjectsAndRefs);

    InterlockedDecrement(&(pMyObj->refsCount));
    long int nRefs = pMyObj->refsCount;

    if (nRefs == 0){
        if (pMyObj->itemNames != NULL){
            GlobalFree(pMyObj->itemNames);
            pMyObj->itemNames = NULL;
            pMyObj->nItems = 0;
        }
        GlobalFree(pMyObj);
    }

    return nRefs;
}

HRESULT STDMETHODCALLTYPE myObj_QueryInterface(IMyObj* pMyObj, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }
    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pMyObj;
        myObj_AddRef(pMyObj);
        return S_OK;
    }

    if (IsEqualGUID(requestedIID, &IID_IContextMenu)){
        *ppv = &(pMyObj->contextMenuImpl);
        myObj_AddRef(pMyObj);

        return S_OK;
    }

    if (IsEqualGUID(requestedIID, &IID_IShellExtInit)){
        *ppv = &(pMyObj->shellExtInitImpl);
        myObj_AddRef(pMyObj);

        return S_OK;
    }


    return E_NOINTERFACE;
}

static IUnknownVtbl IMyObjVtbl = {
    .QueryInterface = &myObj_QueryInterface,
    .AddRef = &myObj_AddRef,
    .Release = &myObj_Release
};



/************** IClassFactory *****************************/
HRESULT STDMETHODCALLTYPE classCreateInstance(IClassFactory* pClassFactory, IUnknown* punkOuter, REFIID pRequestedIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

   // Assume an error by clearing caller's handle.
   *ppv = 0;

   // We don't support aggregation in IExample.
   if (punkOuter){
       return CLASS_E_NOAGGREGATION;
   }

   if (IsEqualGUID(pRequestedIID, &IID_IUnknown) 
       || IsEqualGUID(pRequestedIID, &IID_IContextMenu)
       || IsEqualGUID(pRequestedIID, &IID_IShellExtInit)
       ){
       IMyObj* pMyObj = GlobalAlloc(GPTR, sizeof(IMyObj)); //Zero the memory
       if (pMyObj == NULL){
           return E_OUTOFMEMORY;
       }
   
       pMyObj->lpVtbl = &IMyObjVtbl;

       pMyObj->contextMenuImpl.lpVtbl = &IMyIContextMenuVtbl;
       pMyObj->contextMenuImpl.pBase = pMyObj;

       pMyObj->shellExtInitImpl.lpVtbl = &IMyIShellExtVtbl;
       pMyObj->shellExtInitImpl.pBase = pMyObj;

       return myObj_QueryInterface(pMyObj, pRequestedIID, ppv);
   }

   return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE classAddRef(IClassFactory* pClassFactory){
   return 1;
}

ULONG STDMETHODCALLTYPE classRelease(IClassFactory* pClassFactory){
   return 1;
}

long int nLocks = 0;

HRESULT STDMETHODCALLTYPE classLockServer(IClassFactory* pClassFactory, BOOL flock){
    if (flock){
        InterlockedIncrement(&nLocks);
    } else {
        InterlockedDecrement(&nLocks);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE classQueryInterface(IClassFactory* pClassFactory, REFIID requestedIID, void **ppv){
   // Check if the GUID matches an IClassFactory or IUnknown GUID.
   if (! IsEqualGUID(requestedIID, &IID_IUnknown) && 
       ! IsEqualGUID(requestedIID, &IID_IClassFactory)){
      // It doesn't. Clear his handle, and return E_NOINTERFACE.
      *ppv = 0;
      return E_NOINTERFACE;
   }

   *ppv = pClassFactory;
    pClassFactory->lpVtbl->AddRef(pClassFactory);
   return S_OK;
}

static const IClassFactoryVtbl IClassFactory_Vtbl = {
    classQueryInterface,
    classAddRef,
    classRelease,
    classCreateInstance,
    classLockServer
};

//since we need only one ClassFactory ever:
static IClassFactory IClassFactoryObj = {
    .lpVtbl = &IClassFactory_Vtbl
};


/************************** DLL STUFF ****************************************/

/*
  COM subsystem will call this function in attempt to retrieve an implementation of an IID.
  Usually it's IClassFactory that is requested (but could be anything else?)
 */
__declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID pCLSID, REFIID pIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;
    if (! IsEqualGUID(pCLSID, &SERVER_GUID)){
        //that's not our CLSID, ignore
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    if (! IsEqualGUID(pIID, &IID_IClassFactory)){
        //we only providing ClassFactories in this function
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    //alright, return our implementation of IClassFactory
    *ppv = &IClassFactoryObj;
    return S_OK;
}

__declspec(dllexport) HRESULT PASCAL DllCanUnloadNow(){
    if (nLocks == 0 && nObjectsAndRefs == 0){
        return S_OK;
    }
    return S_FALSE;
}

bool WINAPI entry_point(HINSTANCE hInst, DWORD reason, void* pReserved){
    return true;
}
