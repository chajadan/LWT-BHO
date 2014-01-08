#include "Chunc_CachePageWords.h"
#include "LwtBho.h"

void Chunc_CachePageWords::operator()()
{
	ChuncVar<IHTMLDocument2*> pDoc(nullptr);
	pDoc.AddBound([](IHTMLDocument2* x){return !x;}, false);
	ChuncVar<HRESULT> hr = CoGetInterfaceAndReleaseStream(mpDocStream, IID_IHTMLDocument2, reinterpret_cast<LPVOID*>(pDoc.Get()));
	hr.AddBound([](HRESULT x){return !FAILED(x);});
	chaj::COM::SmartCOMRelease scDoc(pDoc, true); // AddRef and schedule Release

	for (; (int)mi < msp_vWords.Get()->size(); mi += MAX_MYSQL_INLIST_LEN)
	{
		std::list<std::wstring> cWords;
		std::wstring wInList;

		// create buffer of uncached terms
		for (int j = 0; j < MAX_MYSQL_INLIST_LEN && mi+j < msp_vWords.Get()->size(); ++j)
		{
			cache_it it = mCaller->cache->find((*(msp_vWords.Get()))[mi+j]);
			if (mCaller->TermNotCached(it))
			{
				cWords.push_back((*(msp_vWords.Get()))[mi+j]);
				if (!wInList.empty())
					wInList.append(L",");

				wInList.append(L"'");
				wInList.append(mCaller->EscapeSQLQueryValue((*(msp_vWords.Get()))[mi+j]));
				wInList.append(L"'");
			}
		}

		// cache uncached tracked terms
		if (!wInList.empty())
		{
			std::wstring wQuery;
			wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from ");
			wQuery.append(mCaller->wTableSetPrefix);
			wQuery.append(L"words where WoLgID = ");
			wQuery.append(mCaller->wLgID);
			wQuery.append(L" AND WoTextLC in (");
			wQuery.append(wInList);
			wQuery.append(L")");

			EnterCriticalSection(&mCaller->CS_UseDBConn);
			mCaller->DoExecuteDirect(_T("1612isjdlfij"), wQuery);
			while (mCaller->tStmt.fetch_next())
			{
				std::wstring wLC = mCaller->tStmt.field(1).as_string();
				TermRecord rec(mCaller->tStmt.field(2).as_string());
				rec.wRomanization = mCaller->tStmt.field(3).as_string();
				rec.wTranslation = mCaller->tStmt.field(4).as_string();
				mCaller->cache->insert(std::unordered_map<std::wstring,TermRecord>::value_type(wLC, rec));
				cWords.remove(wLC);
			}
			mCaller->tStmt.free_results();
			LeaveCriticalSection(&mCaller->CS_UseDBConn);
		}
		// at this point all terms in list that are tracked are cached

		for (std::wstring& word : cWords)
			mCaller->cache->insert(std::unordered_map<std::wstring,TermRecord>::value_type(word, TermRecord(L"0")));

		if (mCaller->bShuttingDown)
			return;
	}
}