#define _WIN32_WINNT _WIN32_WINNT_WIN7
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

typedef struct {
    IUnknownVtbl* lpVtbl;
    long int refsCount;
} IMyObj;

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

HRESULT STDMETHODCALLTYPE
myObj_QueryInterface(IMyObj* pMyObj, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }
    *ppv = NULL;

    if (! IsEqualGUID(requestedIID, &IID_IUnknown)){
        return E_NOINTERFACE;
    }

    *ppv = pMyObj;
    pMyObj->lpVtbl->AddRef(pMyObj);
    return S_OK;
}

static IUnknownVtbl IMyObjVtbl = {
    .QueryInterface = &myObj_QueryInterface,
    .AddRef = &myObj_AddRef,
    .Release = &myObj_Release
};

HRESULT STDMETHODCALLTYPE classCreateInstance(
        IClassFactory* pClassFactory, IUnknown* punkOuter, REFIID pRequestedIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

   // Assume an error by clearing caller's handle.
   *ppv = 0;

   // We don't support aggregation in IExample.
   if (punkOuter){
       return CLASS_E_NOAGGREGATION;
   }

   if (! IsEqualGUID(pRequestedIID, &IID_IUnknown)){
       return E_NOINTERFACE;
   }

   IMyObj* pMyObj = GlobalAlloc(GMEM_FIXED, sizeof(IMyObj));
   if (pMyObj == NULL){
       return E_OUTOFMEMORY;
   }
   
   pMyObj->lpVtbl = &IMyObjVtbl;
   pMyObj->refsCount = 0;
   pMyObj->lpVtbl->AddRef(pMyObj);


   *ppv = pMyObj;
   return S_OK;
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

