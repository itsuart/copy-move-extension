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

/* GLOBALS */
#include "common.c"

static bool is_server_guid(GUID* test){
    return IsEqualGUID(test, &SERVER_GUID);
}

/* FORWARD DECLARATION EVERYTHING */
struct tag_IMyObj;

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


ULONG STDMETHODCALLTYPE myObj_AddRef(IMyObj* pMyObj);
ULONG STDMETHODCALLTYPE myObj_Release(IMyObj* pMyObj);
HRESULT STDMETHODCALLTYPE myObj_QueryInterface(IMyObj* pMyObj, REFIID requestedIID, void **ppv);

/* BODY */
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

    return S_OK;
}

static IShellExtInitVtbl IMyIShellExtVtbl = {
    .AddRef = &myIShellExtInit_AddRef,
    .Release = &myIShellExtInit_Release,
    .QueryInterface = &myIShellExtInit_QueryInterface,
    .Initialize = &myIShellExtInit_Initialize
};



HRESULT STDMETHODCALLTYPE myIContextMenuImpl_GetCommandString(MyIContextMenuImpl* pImpl, 
                                                              UINT_PTR idCmd, UINT uFlags, UINT* pwReserved,LPSTR pszName,UINT cchMax){
    //not really in use since Vista and we don't provide canonical verbs
    return S_OK;
}

#define COPY_TO_MENU_OFFSET 1
#define MOVE_TO_MENU_OFFSET 2

UINT_PTR CALLBACK OFNHookProc(HWND hdlg, UINT uiMsg, WPARAM wParam, LPARAM lParam){
    if (uiMsg == WM_NOTIFY){
        OFNOTIFYW* pNotify = (OFNOTIFYW*)lParam;
        if (pNotify->hdr.code == CDN_FILEOK){
            return 0;
        }
    }
    return 1;
}


static void cleanup_names_storage(IMyObj* pBase){
    if (pBase->itemNames != NULL){
        GlobalFree(pBase->itemNames);
    }
    pBase->itemNames = NULL;
    pBase->nItems = 0;
}

//I really hate person that causes me "undefined reference to `___chkstk_ms'" errors. And extra calls to clear the memory.
static u16 targetDir[MAX_UNICODE_PATH_LENGTH + 2]; //extra 1 for trailing zero
HRESULT STDMETHODCALLTYPE myIContextMenuImpl_InvokeCommand(MyIContextMenuImpl* pImpl, LPCMINVOKECOMMANDINFO pCommandInfo){
    IMyObj* pBase = pImpl->pBase;
    
    if (pBase->nItems == 0){
        return S_OK; //instantly copy or move nothing!
    }

    const char* pVerb = pCommandInfo->lpVerb;

    if (! IS_INTRESOURCE(pVerb)){
        cleanup_names_storage(pBase);
        return E_FAIL;
    }

    u16* title = NULL;
    uint fileOp;
    
    if ((void*)pVerb == (void*)MAKEINTRESOURCE(COPY_TO_MENU_OFFSET)){
        title  = pBase->nItems > 0 ? L"COPY selected items to:" : L"COPY selected item to:";
        fileOp = FO_COPY;
    } else if ((void*)pVerb == (void*)MAKEINTRESOURCE(MOVE_TO_MENU_OFFSET)){
        title = pBase->nItems > 0 ? L"MOVE selected items to:" : L"MOVE selected item to:";
        fileOp = FO_MOVE;
    } else {
        cleanup_names_storage(pBase);
        return E_FAIL;
    }

    //I want my monads in C!
    HRESULT hr = S_OK; 
    // Create a new common open file dialog.
    IFileOpenDialog *pfd = NULL;
    hr = CoCreateInstance(&CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileOpenDialog, (void**)&pfd);
    if (SUCCEEDED(hr)){
        // Set the dialog as a folder picker.
        DWORD dwOptions;
        hr = pfd->lpVtbl->GetOptions(pfd, &dwOptions);
        if (SUCCEEDED(hr)){
            hr = pfd->lpVtbl->SetOptions(pfd, dwOptions | FOS_PICKFOLDERS);
        }

        // Set the title of the dialog.
        if (SUCCEEDED(hr)){
            hr = pfd->lpVtbl->SetTitle(pfd, title);
        }

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
                        //StringCchCopy(targetDir, sizeof(targetDir)/sizeof(targetDir[0]), pszPath);
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

    pfd->lpVtbl->Release(pfd);
    cleanup_names_storage(pBase);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryContextMenu(MyIContextMenuImpl* pImpl, 
                                                           HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags){
    if (CMF_DEFAULTONLY == (CMF_DEFAULTONLY & uFlags)){
        //show only default menu items, so, non of ours
        //TODO: free memory for names? since we don't even display our menu items
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //in highly unlikely case where is no room left in the menu:
    if (idCmdFirst + 2 > idCmdLast){
        //we don't add any menu enries of ours
        //TODO: free memory for names? since we don't even display our menu items
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //TODO: it might be wiser to extract filenames here

    //insert our menu items: separator, copy and move in that order.
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
    
    u16* copyMenuItemText = pImpl->pBase->nItems > 1 ? L"COPY Selected Items To..." : L"COPY Selected Item To...";
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + COPY_TO_MENU_OFFSET, copyMenuItemText);

    u16* moveMenuItemText = pImpl->pBase->nItems > 1 ? L"MOVE Selected Items To..." : L"MOVE Selected Item To...";
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + MOVE_TO_MENU_OFFSET, moveMenuItemText);

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, (MOVE_TO_MENU_OFFSET + 1)); //shouldn't this be +idCmdFirst?
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

long int nObjectsAndRefs = 0;

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

static IContextMenuVtbl IMyIContextMenuVtbl = {
    .QueryInterface = &myIContextMenuImpl_QueryInterface,
    .AddRef = &myIContextMenuImpl_AddRef,
    .Release = &myIContextMenuImpl_Release,
    .GetCommandString = &myIContextMenuImpl_GetCommandString,
    .QueryContextMenu = &myIContextMenuImpl_QueryContextMenu,
    .InvokeCommand = &myIContextMenuImpl_InvokeCommand
};


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

HRESULT STDMETHODCALLTYPE
classQueryInterface(IClassFactory* pClassFactory, REFIID requestedIID, void **ppv){
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
    &IClassFactory_Vtbl
};

/*
  COM subsystem will call this function in attempt to retrieve an implementation of an IID.
  Usually it's IClassFactory that is requested (but could be anything else?)
 */
__declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID pCLSID, REFIID pIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;
    if (! is_server_guid(pCLSID)){
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
