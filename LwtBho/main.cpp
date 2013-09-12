#include "LwtBho.h"

#ifndef _DEBUG
	const CLSID BhoCLSID = {0xe72bb92,0x73d4,0x4bef,0xbc,0x8,0xfe,0x3b,0x96,0x85,0x93,0xa3};
	#define BhoCLSIDs  L"{0E72BB92-73D4-4BEF-BC08-FE3B968593A3}"
	static const GUID CLSID_AddObject = 
	{ 0xe72bb92, 0x73d4, 0x4bef, { 0xbc, 0x8, 0xfe, 0x3b, 0x96, 0x85, 0x93, 0xa3} };
#else
	const CLSID BhoCLSID = {0x2bd807b2,0xb9a5,0x44f3,0x9a,0x32,0x7a,0x59,0xb0,0xdb,0xc8,0x55};
	#define BhoCLSIDs L"{2BD807B2-B9A5-44F3-9A32-7A59B0DBC855}"
	static const GUID CLSID_AddObject = 
	{ 0x2bd807b2, 0xb9a5, 0x44f3, { 0x9a, 0x32, 0x7a, 0x59, 0xb0, 0xdb, 0xc8, 0x55 } };
#endif

class MyClassFactory : public IClassFactory
{
	long ref;
public:
	
	// IUnknown... (nb. this class is instantiated statically, which is why Release() doesn't delete it.)
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv) {if (riid==IID_IUnknown || riid==IID_IClassFactory) {*ppv=this; AddRef(); return S_OK;} else return E_NOINTERFACE;}
	ULONG STDMETHODCALLTYPE AddRef() {InterlockedIncrement(&gref); return InterlockedIncrement(&ref);}
	ULONG STDMETHODCALLTYPE Release() {int tmp = InterlockedDecrement(&ref); InterlockedDecrement(&gref); return tmp;}
  
	// IClassFactory...
	HRESULT STDMETHODCALLTYPE LockServer(BOOL b) {if (b) InterlockedIncrement(&gref); else InterlockedDecrement(&gref); return S_OK;}
	HRESULT STDMETHODCALLTYPE CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, LPVOID *ppvObj)
	{
		*ppvObj = NULL;
		if (pUnkOuter)
			return CLASS_E_NOAGGREGATION;
		
		LwtBho *bho=new LwtBho();
		bho->AddRef();
		HRESULT hr=bho->QueryInterface(riid, ppvObj);
		bho->Release();
		return hr;
	}
  
	// MyClassFactory...
	MyClassFactory() : ref(0) {}
};


STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppvOut)
{
	static MyClassFactory factory; *ppvOut = NULL;
	if (rclsid==BhoCLSID) {return factory.QueryInterface(riid,ppvOut);}
	else return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow(void)
{
	return (gref>0) ? S_FALSE : S_OK;
}

wstring GetDllPath()
{
#ifdef _DEBUG
	wstring strMod = L"LwtBho_d.dll";
#else
	wstring strMod = L"LwtBho.dll";
#endif

	HMODULE hMod = GetModuleHandle(strMod.c_str());
	if (!hMod)
		return L"";

	wchar_t szBuff[MAX_PATH] = L"";
	DWORD dwRet = GetModuleFileName(hMod, (LPWSTR)szBuff, sizeof(szBuff));
	if (!dwRet || GetLastError() == ERROR_INSUFFICIENT_BUFFER)
		return L"";

	return wstring(szBuff);
}

HRESULT __stdcall DllRegisterServer(void)
{
	HKEY pKey;

	wstring wDllPath = GetDllPath();
	wstring wClsid = L"CLSID\\";
	wstring wClsEntry = wClsid += BhoCLSIDs;
	wstring wClsInproc = wClsEntry += L"\\InprocServer32";
	wstring wBHOKey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects\\";
	wstring wAddonName = L"Learning With Texts Extension";
#ifdef _DEBUG
		wAddonName += L" (Debug)";
#endif
	wstring wRes = L"res://";	wRes += wDllPath;	wRes.append(L"/");	wRes.append(to_wstring(IDR_LWTJAVASCRIPT02));
	
	DWORD dwContexts = 0xFFFFFFFF;

	if (wDllPath.size() == 0) return E_UNEXPECTED;

	chaj::util::CatchSentinel<long> cs(ERROR_SUCCESS, true);
	cs += RegCreateKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\MenuExt\\Lwt Settings\\", NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &pKey, NULL);
	cs += RegSetValueEx(pKey, NULL, NULL, REG_SZ, (BYTE*)wRes.c_str(), wRes.length()*sizeof(wstring::value_type)+2);
	cs += RegSetValueEx(pKey, L"Contexts", NULL, REG_DWORD, (BYTE*)&dwContexts, sizeof(DWORD));
	cs += RegCreateKeyEx(HKEY_CLASSES_ROOT, wClsEntry.c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &pKey, NULL);
	cs += RegSetValueEx(pKey, NULL, NULL, REG_SZ, (BYTE*)wAddonName.c_str(), wAddonName.length()*sizeof(wstring::value_type)+2);
	cs += RegCreateKeyEx(HKEY_CLASSES_ROOT, wClsInproc.c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &pKey, NULL);
	cs += RegSetValueEx(pKey, NULL, NULL, REG_SZ, (BYTE*)wDllPath.c_str(), wDllPath.length()*sizeof(wstring::value_type)+2);
	cs += RegCreateKeyEx(HKEY_LOCAL_MACHINE, (wBHOKey += BhoCLSIDs).c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &pKey, NULL);

	if (cs.Seen())
	{
		DllUnregisterServer();
		PrintFormattedError(cs.LastException());
		return S_FALSE;
	}
	else
		return S_OK;
}

HRESULT __stdcall DllUnregisterServer()
{
	wstring wClsid = L"CLSID\\";
	wstring wClsEntry = wClsid += wstring(BhoCLSIDs); // CLSID\\{0E72BB92-73D4-4BEF-BC08-FE3B968593A3}
	wstring wInproc = L"\\InprocServer32";
	wstring wBHO = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects\\";

	chaj::util::CatchSentinel<long> cs(ERROR_SUCCESS, true);
	cs += RegDeleteKeyEx(HKEY_LOCAL_MACHINE, (wBHO += wstring(BhoCLSIDs)).c_str(), KEY_WOW64_32KEY, NULL);
	cs += RegDeleteKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\MenuExt\\Lwt Settings\\", KEY_WOW64_32KEY, NULL);
	cs += RegDeleteKeyEx(HKEY_CLASSES_ROOT, (wClsEntry += wInproc).c_str(), KEY_WOW64_32KEY, NULL);
	cs += RegDeleteKeyEx(HKEY_CLASSES_ROOT, wClsEntry.c_str(), KEY_WOW64_32KEY, NULL);

	if (cs.Seen())
	{
		PrintFormattedError(cs.LastException());
		return S_FALSE;
	}
	else
		return S_OK;
}

extern "C" BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	hInstance = hinstDLL;
	switch(fdwReason)
	{
		case DLL_PROCESS_ATTACH:
		{
			DisableThreadLibraryCalls(hinstDLL);

			TCHAR pszLoader[MAX_PATH];
			GetModuleFileName(NULL, pszLoader, MAX_PATH);
			_tcslwr(pszLoader);
			if (_tcsstr(pszLoader, _T("explorer.exe"))) 
				return FALSE;

			return TRUE;
		}
	}
}