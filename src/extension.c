#define _WIN32_WINNT _WIN32_WINNT_WIN7
#define UNICODE

#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>
#include <Shobjidl.h>

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;

/* GLOBALS */
#include "common.c"

static bool is_server_guid(GUID* test){
    return IsEqualGUID(test, &SERVER_GUID);
}
/*
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
*/
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

HRESULT STDMETHODCALLTYPE myIShellExtInit_Initialize(MyIShellExtInitImpl* pImpl, LPCITEMIDLIST pIDFolder, IDataObject *pDataObj, HKEY hRegKey){
    //TODO: implement
    return S_OK;
}

static IShellExtInitVtbl IMyIShellExtVtbl = {
    .AddRef = &myIShellExtInit_AddRef,
    .Release = &myIShellExtInit_Release,
    .QueryInterface = &myIShellExtInit_QueryInterface,
    .Initialize = &myIShellExtInit_Initialize
};



HRESULT STDMETHODCALLTYPE myIContextMenuImpl_GetCommandString(
MyIContextMenuImpl* pImpl, UINT_PTR idCmd, UINT uFlags, UINT* pwReserved,LPSTR pszName,UINT cchMax){
     CreateFileW( L"C:\\tmp\\GetCommandString",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    //not really in use since Vista and we don't provide canonical verbs
    return S_OK;
}

#define COPY_TO_MENU_OFFSET 1
#define MOVE_TO_MENU_OFFSET 2

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_InvokeCommand(MyIContextMenuImpl* pImpl, LPCMINVOKECOMMANDINFO pCommandInfo){
     CreateFileW( L"C:\\tmp\\InvokeCommand",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    /* MSDN:
If a canonical verb exists and a menu handler does not implement the canonical verb, it must return a failure code to enable the next handler to be able to handle this verb. Failing to do this will break functionality in the system including ShellExecute.

Alternatively, rather than a pointer, this parameter can be MAKEINTRESOURCE(offset) where offset is the menu-identifier offset of the command to carry out. Implementations can use the IS_INTRESOURCE macro to detect that this alternative is being employed. The Shell uses this alternative when the user chooses a menu command.

    */
    //Since we do not provide verbs, just return failure for all non-offset argument
    const char* pVerb = pCommandInfo->lpVerb;
    if (! IS_INTRESOURCE(pVerb)){
        return E_FAIL;
    }

    //TODO: properly handle nShow argument

    //MSDN says it's important to use pCommandInfo->hwnd and using NULL could ruin all the things, so
    if (pVerb == MAKEINTRESOURCE(COPY_TO_MENU_OFFSET)){
        MessageBoxW(pCommandInfo->hwnd, L"You clicked COPY", L"My ext", MB_OK | MB_ICONINFORMATION);
        return S_OK;
    }

    if (pVerb == MAKEINTRESOURCE(MOVE_TO_MENU_OFFSET)){
        MessageBoxW(pCommandInfo->hwnd, L"You clicked MOVE", L"My ext", MB_OK | MB_ICONINFORMATION);
        return S_OK;
    }

    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryContextMenu(MyIContextMenuImpl* pImpl, 
                                                           HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags){
     CreateFileW( L"C:\\tmp\\QueryContextMenu_init",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    if (CMF_DEFAULTONLY == (CMF_DEFAULTONLY & uFlags)){
        //show only default menu items, so, non of ours
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //in highly unlikely case where is no room left in the menu:
    if (idCmdFirst + 2 > idCmdLast){
        //we don't add any menu enries of ours
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //insert our menu items: separator, copy and move in that order.
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + COPY_TO_MENU_OFFSET, L"COPY Selected Items To...");
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + MOVE_TO_MENU_OFFSET, L"MOVE Selected Items To...");

     CreateFileW( L"C:\\tmp\\QueryContextMenu_success",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

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
     CreateFileW( L"C:\\tmp\\query_IContextMenu_success",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

        return S_OK;
    }

    if (IsEqualGUID(requestedIID, &IID_IShellExtInit)){
        *ppv = &(pMyObj->shellExtInitImpl);
        myObj_AddRef(pMyObj);
     CreateFileW( L"C:\\tmp\\query_ISHellExtInit_success",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

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
       IMyObj* pMyObj = GlobalAlloc(GMEM_FIXED, sizeof(IMyObj));
       if (pMyObj == NULL){
           return E_OUTOFMEMORY;
       }
   
       pMyObj->lpVtbl = &IMyObjVtbl;
       pMyObj->refsCount = 0;
       
       pMyObj->contextMenuImpl.lpVtbl = &IMyIContextMenuVtbl;
       pMyObj->contextMenuImpl.pBase = pMyObj;

       pMyObj->shellExtInitImpl.lpVtbl = &IMyIShellExtVtbl;
       pMyObj->shellExtInitImpl.pBase = pMyObj;

     CreateFileW( L"C:\\tmp\\classCreateInstance_success",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);


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

   // First, we fill in his handle with the same object pointer he passed us.
   // That's our IClassFactory (MyIClassFactoryObj) he obtained from us.
   *ppv = pClassFactory;

   // Call our IClassFactory's AddRef, passing the IClassFactory. 
   pClassFactory->lpVtbl->AddRef(pClassFactory);

     CreateFileW( L"C:\\tmp\\classfactory_query_success",
                            GENERIC_WRITE,
                            0,
                            NULL,
                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);


   // Let him know he indeed has an IClassFactory. 
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
