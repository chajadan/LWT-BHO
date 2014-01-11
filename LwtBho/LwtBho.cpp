#include "LwtBho.h"

LONG gref = 0;
wstring wStatIntro(L"lwtStat");
HINSTANCE hInstance;

using namespace chaj::DOM;

 DOMIteratorFilter* LwtBho::s_pFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTerm);
 DOMIteratorFilter* LwtBho::s_pFullFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTermAndSent);

 void LwtBho::ConnectToDatabase()
 {
	if (!tConn.lwtConnect(hBrowser))
		throw new exception;
 }
void LwtBho::DeleteCriticalSections()
{
	DeleteCriticalSection(&CS_UseDBConn);
	DeleteCriticalSection(&CS_General);
}
void LwtBho::GetLwtActiveTableSetPrefix()
{
	if (tStmt.execute_direct(tConn, L"select COALESCE(LWTValue, '') from _lwtgeneral where LWTKey = 'current_table_prefix'"))
	{
		if (tStmt.fetch_next())
		{
			wTableSetPrefix = tStmt.field(1).as_string();
			if (wTableSetPrefix.size()) // only a non-null prefix should be _ terminated, the default is not such
				wTableSetPrefix.append(L"_");
		}
		tStmt.free_results();
	}
}
void LwtBho::InitializeCriticalSections()
{
	if (!InitializeCriticalSectionAndSpinCount(&CS_UseDBConn, 0x00000400))
	{
		TRACE(L"Cound not initialize critical section CS_UseDBConn in LwtBho. 5110suhdg\n");
		throw new exception;
	}
	if (!InitializeCriticalSectionAndSpinCount(&CS_General, 0x00000400))
	{
		TRACE(L"Cound not initialize critical section CS_General in LwtBho. lhjdsf87e\n");
		throw new exception;
	}
}
void LwtBho::LoadCssFile()
{
	HRSRC rc = ::FindResource(hInstance, MAKEINTRESOURCE(IDR_LWTCSS), MAKEINTRESOURCE(CSS));
	HGLOBAL rcData = ::LoadResource(hInstance, rc);
	wCss = to_wstring(static_cast<const char*>(::LockResource(rcData)));
	assert(wCss.size());
}
void LwtBho::LoadJavascriptFile()
{
	HRSRC rc = ::FindResource(hInstance, MAKEINTRESOURCE(IDR_LWTJAVASCRIPT01), MAKEINTRESOURCE(JAVASCRIPT));
	HGLOBAL rcData = ::LoadResource(hInstance, rc);
	wJavascript = ::to_wstring(static_cast<const char*>(::LockResource(rcData)));
	assert(wJavascript.size());
}
void LwtBho::WaitForThreadsToTerminate()
{
	// wait for all threads to shut down
	// thread is not deleted as this was causing issues for unknown reasons (would need troubleshooting)
	// LwtBho destructor will free all objects soon enough anyway
	for (unsigned int i = 0; i < cpThreads.size(); ++i)
	{
		if (cpThreads[i]->joinable())
			cpThreads[i]->join();
	}

	int sleepLength = 0;
	while (mNumDetachedThreads > 0)
	{
		Sleep(1 * 1000);
		++sleepLength;
		if (sleepLength > 60)
			TRACE(L"In LwtBho desctructor, active threads still remain after shutdown timeout.\n");
	}
}