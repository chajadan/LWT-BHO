#ifndef LwtBho_INCLUDED
#define LwtBho_INCLUDED

#include "LWTCache.h"
#include "Resource/resource.h"
#include "assert.h"
#include "dbdriver.h"
#include <Windows.h>
#include <thread>
#include "tchar.h"
#include "stdio.h"
#include "Ocidl.h"
#include "exdisp.h"
#include "exdispid.h"
#include "shlguid.h"
#include "mshtml.h"
#include "comutil.h"
#include "connection.h"
#include "query.h"
#include "result.h"
#include "manip.h"
#include <fstream>
#include <sstream>
#include <atomic>
#include <condition_variable>
#include <vector>
#include <map>
#include <unordered_map>
#include <unordered_set>
#include <set>
#include <forward_list>
#include <stack>
#include <algorithm>
#include <regex>
#include <mshtmdid.h>
#include "comip.h"
#include "tiodbc.hpp"
#include "chajUtil.h"
#include "chajDOM.h"
#include "Chunc_CachePageWords.h"
using namespace mysqlpp;
using namespace std;
using namespace chaj;
using namespace chaj::DOM;
using namespace chaj::COM;
using chaj::str::to_wstring;

typedef regex_iterator<wstring::const_iterator, wchar_t,regex_traits<wchar_t>> wregex_iterator;
typedef vector<wstring>::const_iterator wstr_c_it;
const INT_PTR CTRL_FORCE_PARSE = static_cast<INT_PTR>(101);
const INT_PTR CTRL_CHANGE_LANG = static_cast<INT_PTR>(102);
const int MWSpanAltroSize = 7;
const int MAX_FIELD_WIDTH = 255;
const wstring wstrNewline(L"&#13;&#10;");

#define nHiddenChunkOpenBookmarkLen 3
#define LWT_MAXWORDLEN 255
#define LWT_MAX_MWTERM_LENGTH 9
extern wstring wHiddenChunkBookmark;
extern wstring wPTagBookmark;
extern wstring wStatIntro;
extern HINSTANCE hInstance;
extern vector<HANDLE> vDelete;
extern LONG gref;

struct DlgResult
{
	wstring wstrTerm;
	int nStatus;
};

struct MainDlgStruct
{
	MainDlgStruct(bool bIsLinkPart = false) : bOnLink(bIsLinkPart) {};
	vector<wstring> Terms;
	bool bOnLink;
};

class Token_Word
{
public:
	std::wstring wToken;
	bool bDigested;
};
class Token_Sentence
{
public:
	Token_Sentence(bool bIsDigested) : bDigested(bIsDigested) {}
	void push_back(wstring wWord, bool bIsDigested = true) {Token_Word tw; tw.wToken = wWord; tw.bDigested = bIsDigested; vWords.push_back(tw);}
	const wstring& operator[](int i) const {return vWords[i].wToken;}
	vector<Token_Word> vWords;
	vector<Token_Word>::size_type size() const {return vWords.size();}
	bool isWordDigested(int indWord) const {return vWords[indWord].bDigested;}
	bool bDigested;
};
class TokenStruct
{
public:
	const Token_Sentence& operator[](int i) const {return vSentences[i];}
	vector<Token_Sentence> vSentences;
	void push_back(Token_Sentence ts) {vSentences.push_back(ts);}
	void push_back(wstring wNonDigested) {Token_Sentence ts(false); ts.push_back(wNonDigested, false); vSentences.push_back(ts);}
	vector<Token_Sentence>::size_type size() const {return vSentences.size();}
};

class LwtBho: public IObjectWithSite, public IDispatch
{
	friend class LWTCache;
	friend class Chunc_CachePageWords;

private:

	struct mwVals
	{
		mwVals(wstring wStat, wstring wWC, wstring wTerm) : wStatus(wStat), wWordCount(wWC), wMWTerm(wTerm) {};
		mwVals(const TermRecord* RecPtr, wstring wWC, wstring wTerm) : pRec(RecPtr), wWordCount(wWC), wMWTerm(wTerm) {};
		const TermRecord* pRec;
		wstring wStatus;
		wstring wWordCount;
		wstring wMWTerm;
	};

public:
	LwtBho();
	~LwtBho();

	void LoadCssFile()
	{
		HRSRC rc = ::FindResource(hInstance, MAKEINTRESOURCE(IDR_LWTCSS),
			MAKEINTRESOURCE(CSS));
		HGLOBAL rcData = ::LoadResource(hInstance, rc);
		wCss = to_wstring(static_cast<const char*>(::LockResource(rcData)));
	}
	void test3()
	{
		//DBDriver* dbdriv = conn.driver();
		//dbdriv->client_version();
		//MessageBoxA(NULL, dbdriv->client_version().c_str(), "dcversion", MB_OK);
		//SetCharsetNameOption scno("utf8");

		mb("option set");
		//st_mysql_options smo = dbdriv->get_options();
		//mb("charset");
		//mb(smo.charset_name);
		Query q = conn.query("select LgRegexpWordCharacters from languages where LgID = 3");
		StoreQueryResult sqr = q.store();
		wchar_t res[200] = L"";
		//ToUCS2(res, 200, sqr[0][0]);
			mb("the result!!");
		mb(res);
	}
	void DoError(const tstring& errCode, const tstring& errMsg)
	{
		mb(errMsg, errCode);
	}
	bool DoExecute(const wstring& errCode)
	{
		bool success = tStmt.execute();
		if (!success)
			DoError(errCode, tStmt.last_error());
		return success;
	}
	bool DoExecuteDirect(const wstring& errCode, const wstring& tSql)
	{
		tStmt.free_results(); // just in case
		bool success = tStmt.execute_direct(tConn, tSql);
		if (!success)
		{
			wstring werr = tStmt.last_error();
			TRACE(L"execute_direct error: %s %s\n", errCode.c_str(), werr.c_str());
		}
		return success;
	}
	void AddNewTerm(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		wstring wTermEscaped = EscapeSQLQueryValue(wTerm);
		wstring wQuery;
		wQuery.append(L"insert into ");
		wQuery.append(wTableSetPrefix);
		wQuery.append(L"words (WoLgID, WoText, WoTextLC, WoStatus, WoStatusChanged,  WoTodayScore, WoTomorrowScore, WoRandom) values (");
		wQuery.append(wLgID);
		wQuery.append(L", '");
		wQuery.append(wTermEscaped);
		wQuery.append(L"', '");
		wQuery.append(wTermEscaped);
		wQuery.append(L"', ");
		wQuery.append(wNewStatus);
		wQuery.append(L", NOW(), CASE WHEN WoStatus > 5 THEN 100 ELSE (((POWER(2.4,WoStatus) + WoStatus - DATEDIFF(NOW(),WoStatusChanged) - 1) / WoStatus - 2.4) / 0.14325248) END, CASE WHEN WoStatus > 5 THEN 100 ELSE (((POWER(2.4,WoStatus) + WoStatus - DATEDIFF(NOW(),WoStatusChanged) - 2) / WoStatus - 2.4) / 0.14325248) END, RAND())");
		if (tStmt.execute_direct(tConn, wQuery))
		{
			/// value may already be in cache if it was untracked in the scope of the current cache
			InsertOrUpdateCacheItem(wTerm, wNewStatus);

			if (IsMultiwordTerm(wTerm))
			{
				UpdateCacheMWFragments(wTerm);
				AddNewMWSpans(wTerm, wNewStatus, pDoc);
			}
			else
				UpdatePageSpans(pDoc, wTerm, wNewStatus);
		}
		else
		{
			unordered_map<wstring,TermRecord>::iterator it = cache->find(wTerm);
			if (it != cache->end())
				ChangeTermStatus(wTerm, wNewStatus, pDoc);
			else
				mb(L"Unable to add term. This is a database related issue. Error code: 325nijok", wTerm);
		}
	}
	void AddNewMWSpans(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		if (!wTerm.size())
			return;

		vector<wstring> vParts;

		if (bWithSpaces)
		{
			int pos = -1; //offset to jive with do loop
			wchar_t* wcsTerm = new wchar_t[wTerm.size() + 1];
			wcscpy(wcsTerm, wTerm.c_str());
			wchar_t* pwc = wcstok(wcsTerm, L" ");
			while (pwc)
			{
				vParts.push_back(pwc);
				pwc = wcstok(NULL, L" ");
			}
			delete [] wcsTerm;
		}
		else
		{
			for (wstring::size_type i = 0; i < wTerm.size(); ++i)
				vParts.push_back(to_wstring(wTerm.c_str()[i]));
		}

		assert(vParts.size());

		HRESULT hr;
		unsigned int uCount = WordsInTerm(wTerm);
		VARIANT_BOOL vBool;
		_bstr_t bFind(wTerm.c_str());
		wstring out;
		unordered_map<wstring,TermRecord>::const_iterator it = cache->cfind(wTerm);
		assert(it != cache->cend());

		IHTMLElement* pBody = chaj::DOM::GetBodyFromDoc(pDoc);
		assert(pBody);
		AppendTermDivRec(pBody, wTerm, it->second);
		pBody->Release();
		AppendMWSpan(out, wTerm, &(it->second), to_wstring(uCount));

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		assert(pRange);

		bool bContinueAfterBreak = false;
		
		do
		{
			hr = pRange->findText(bFind, 0, 0, &vBool);
			assert(SUCCEEDED(hr));

			if (vBool != VARIANT_TRUE)
				break;

			_bstr_t bMWBookmark;
			hr = pRange->getBookmark(bMWBookmark.GetAddress());
			assert(SUCCEEDED(hr));

			IHTMLElement* pStartElem = nullptr;

			for (vector<wstring>::size_type i = 0; i < vParts.size(); ++i)
			{
				_bstr_t bMWTermPart(vParts[i].c_str());
				pRange->findText(bMWTermPart.GetBSTR(), 0, 0, &vBool);
				assert(vBool == VARIANT_TRUE);

				IHTMLElement* pMWPartSpan = nullptr;
				hr = pRange->parentElement(&pMWPartSpan);
				assert(SUCCEEDED(hr));

				while (pMWPartSpan && !chaj::DOM::GetAttributeValue(pMWPartSpan, L"lwtterm").size())
				{
					IHTMLElement* pHigherParent = nullptr;
					hr = pMWPartSpan->get_parentElement(&pHigherParent);
					assert(SUCCEEDED(hr));

					pMWPartSpan->Release();
					pMWPartSpan = pHigherParent;
				}

				if (!pMWPartSpan || GetAttributeValue(pMWPartSpan, L"lwtterm") != vParts[i])
				{
					bContinueAfterBreak = true;
					break;
				}

				if (i == 0)
					pStartElem = pMWPartSpan;
				else
					pMWPartSpan->Release();

				hr = chaj::DOM::TxtRange_CollapseToEnd(pRange);
				assert(SUCCEEDED(hr));
				hr = chaj::DOM::TxtRange_RevertEnd(pRange);
				assert(SUCCEEDED(hr));
			}

			hr = pRange->moveToBookmark(bMWBookmark.GetBSTR(), &vBool);
			assert(SUCCEEDED(hr));

			if (!bContinueAfterBreak) // we found a term to highlight
			{
				hr = chaj::DOM::AppendHTMLBeforeBegin(out, pStartElem);
				assert(SUCCEEDED(hr));
				pStartElem->Release();

				hr = chaj::DOM::TxtRange_CollapseToEnd(pRange);
				assert(SUCCEEDED(hr));

				hr = chaj::DOM::TxtRange_RevertEnd(pRange);
				assert(SUCCEEDED(hr));
			}

			if (pStartElem) pStartElem->Release();

			if (bContinueAfterBreak)
			{
				bContinueAfterBreak = false;

				hr = chaj::DOM::TxtRange_RevertEnd(pRange);
				assert(SUCCEEDED(hr));

				hr = chaj::DOM::TxtRange_MoveStartByChars(pRange, 1);
				assert(SUCCEEDED(hr));

				continue;
			}
		} while (true);

		pRange->Release();
	}
	HRESULT RemoveNodeFromElement(IHTMLElement* pElement)
	{
		IHTMLDOMNode* pNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pElement);
		if (pNode)
		{
			IHTMLDOMNode* pRemoved;
			HRESULT hr = pNode->removeNode(VARIANT_TRUE, &pRemoved);
			pRemoved->Release();
			return hr;
		}
		
		return E_NOINTERFACE;
	}
	void RemoveMWSpans(const wstring& wTerm, IHTMLDocument2* pDoc)
	{
		unsigned int uCount = WordsInTerm(wTerm);
		VARIANT_BOOL vBool;
		_bstr_t bFind(uCount);
		_bstr_t bstrChar(L"character");
		IHTMLElement* pParent = nullptr;
		long lVal;

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		assert(pRange);

		pRange->findText(bFind, 1000000, 0, &vBool);
		
		while (vBool == VARIANT_TRUE)
		{
			pRange->parentElement(&pParent);
			if (pParent)
			{
				wstring wCurTerm = GetAttributeValue(pParent, L"lwtTerm");
				if (wTerm == wCurTerm)
				{
					IHTMLElement* pSup = nullptr;
					pParent->get_parentElement(&pSup);
					assert(pSup);
					HRESULT hr = RemoveNodeFromElement(pSup);
					assert(SUCCEEDED(hr));
#ifdef _DEBUG
					if (FAILED(hr))
						TRACE(L"%s", L"Unable to properly delete multiword span.\n");
#endif
				}
			}

			pRange->move(bstrChar, 1, &lVal); // todo I'm not sure if this could potentially skip a single number, or how it responds to the previous node removal in terms of txtrange cursor position, but I'm hoping, with spaces, one char is reasonable = hack
			pRange->findText(bFind, 1000000, 0, &vBool);
		}
	}
	void RemoveMWSpans2(const wstring& wTerm)
	{
		IDispatch* pRoot = NULL, *pCur = NULL;
		pTW->get_root(&pRoot);
		pTW->putref_currentNode(pRoot);
		pCur = pRoot;

		wstring wCurTerm;

		do
		{
			IHTMLElement* pEl = GetAlternateInterface<IDispatch,IHTMLElement>(pCur);
			if (pEl)
			{
				wCurTerm = GetAttributeValue(pEl, L"lwtTerm");
				pEl->Release();
				if (wTerm == wCurTerm)
				{
					IHTMLDOMNode* pNode = GetAlternateInterface<IDispatch,IHTMLDOMNode>(pCur);
					if (pNode)
					{
						IHTMLDOMNode* pRemoved;
						HRESULT hr = pNode->removeNode(VARIANT_TRUE, &pRemoved);
						pRemoved->Release();
					}
				}
			}
			pTW->nextNode(&pCur);
		} while (pCur);
	}
	void RemoveTerm(IHTMLDocument2* pDoc, const wstring& wTerm)
	{
		wstring wQuery;
		wQuery.append(L"delete from ");
		wQuery.append(wTableSetPrefix);
		wQuery.append(L"words where WoTextLC = '");
		wQuery.append(wTerm);
		wQuery.append(L"' and WoLgID = ");
		wQuery.append(wLgID);
		if (tStmt.execute_direct(tConn, wQuery))
		{
			AlterCacheItemStatus(wTerm, L"0");
			if (IsMultiwordTerm(wTerm))
				RemoveMWSpans(wTerm, pDoc);//ReparsePage(pDoc);
			else
				UpdatePageSpans(pDoc, wTerm, L"0");
		}
		else
			mb(L"Unable to remove term.", wTerm);
	}
	void AlterCacheItemStatus(const wstring& wTerm, const wstring& wNewStatus)
	{
		unordered_map<wstring,TermRecord>::iterator it = cache->find(wTerm);
		assert(it != cache->end());
		(*it).second.wStatus = wNewStatus;
	}
	void InsertOrUpdateCacheItem(const wstring& wTerm, const wstring& wNewStatus)
	{
		cache_it it = cache->find(wTerm);
		if (it == cache->end())
			cache->insert(unordered_map<wstring,TermRecord>::value_type(wTerm, TermRecord(wNewStatus)));
		else
			(*it).second.wStatus = wNewStatus;
	}
	void ChangeTermStatus(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		if (wNewStatus == L"0")
			return RemoveTerm(pDoc, wTerm);

		wstring wQuery;
		wQuery.append(L"update ");
		wQuery.append(wTableSetPrefix);
		wQuery.append(L"words set WoStatus = ");
		wQuery.append(wNewStatus);
		wQuery.append(L", WoStatusChanged = Now(), WoTodayScore = CASE WHEN WoStatus > 5 THEN 100 ELSE (((POWER(2.4,WoStatus) + WoStatus - DATEDIFF(NOW(),WoStatusChanged) - 1) / WoStatus - 2.4) / 0.14325248) END, WoTomorrowScore = CASE WHEN WoStatus > 5 THEN 100 ELSE (((POWER(2.4,WoStatus) + WoStatus - DATEDIFF(NOW(),WoStatusChanged) - 2) / WoStatus - 2.4) / 0.14325248) END, WoRandom = RAND() where WoTextLC = '");
		wQuery.append(wTerm);
		wQuery.append(L"' and WoLgID = ");
		wQuery.append(wLgID);
		if (!DoExecuteDirect(L"292hduhsal", wQuery))
			return;

		AlterCacheItemStatus(wTerm, wNewStatus);
		UpdatePageSpans(pDoc, wTerm, wNewStatus);
	}
	void InvokeScripts(IHTMLDocument2* pDoc)
	{
		IHTMLElement* pBody = GetBodyFromDoc(pDoc);
		IDOMTreeWalker* pTree = GetTreeWalkerWithFilter(pDoc, pBody, &FilterNodes_Scripts, SHOW_ELEMENT);
		pBody->Release();

		if (pTree)
		{

		}
	}
	void  UpdatePageSpans2(IHTMLDocument2* pDoc, const wstring& wTerm, const wstring& wNewStatus)
	{
		IDispatch* pDispNode = NULL;
		pNI->nextNode(&pDispNode);
		IHTMLElement* pEl = GetHTMLElementFromDisp(pDispNode);
		wstring wTerm2 = GetAttributeValue(pEl, L"lwtTerm");
		if (wTerm2 == wTerm)
		int i = 0;
	}
	/*
		WordsInTerm
		input condition: expects some valid term in canonical form
	*/
	unsigned int WordsInTerm(const wstring& wstrTerm)
	{
		if (bWithSpaces == true)
		{
			if (wstrTerm.size() == 0)
				return 0;

			unsigned int count = 1; // at least one word present
		
			for (wstring::size_type i = 0; i < wstrTerm.size(); ++i)
				if (wstrTerm[i] == L' ') // found a(nother) space
					++count; // so additonal word

			return count;
		}
		else
			return wstrTerm.size();
	}
	void UpdatePageSpansMW(IHTMLDocument2* pDoc, const wstring& wTerm, const wstring& wNewStatus)
	{
		wstring out;
		IHTMLElement* pParent = nullptr;
		VARIANT_BOOL vBool;
		long lVal;

		unsigned int uWordCount = WordsInTerm(wTerm);

		wstring wNewStat = L"lwtStat";
		wNewStat.append(wNewStatus);

		_bstr_t bstrClassVal(wNewStat.c_str());
		_bstr_t bstrChar(L"character");
		_bstr_t bWord(wTerm.c_str());
		_bstr_t bWordCount(uWordCount);

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		assert(pRange);

		pRange->findText(bWord, 1000000, 2, &vBool);
		
		while (vBool == VARIANT_TRUE)
		{
			BSTR bsBMark;
			HRESULT hr = pRange->getBookmark(&bsBMark);
			assert(SUCCEEDED(hr));

			pRange->findText(bWordCount, -500, 0, &vBool);
			if (vBool == VARIANT_TRUE)
			{
				pRange->parentElement(&pParent);
				assert(pParent != nullptr);
				wstring wCurTerm = GetAttributeValue(pParent, L"lwtTerm");
				if (wTerm == wCurTerm)
					pParent->put_className(bstrClassVal.GetBSTR());
				pParent->Release();
				//pRange->move(bstrChar, wTerm.size() + MWSpanAltroSize, &lVal); //todo this is a hack to me, but perhaps preferable to duplicating a txtrange
			}
			//else
			
			hr = pRange->moveToBookmark(bsBMark, &vBool);
			assert(SUCCEEDED(hr));
			SysFreeString(bsBMark);
			pRange->move(bstrChar, 1, &lVal);
			pRange->findText(bWord, 1000000, 2, &vBool);
		}
		pRange->Release();
	}
	void UpdatePageSpans4(IHTMLDocument2* pDoc, const wstring& wTerm, const wstring& wNewStatus)
	{
		if (IsMultiwordTerm(wTerm))
			return UpdatePageSpansMW(pDoc, wTerm, wNewStatus);

		wstring wNewStat = L"lwtStat";
		wNewStat.append(wNewStatus);
		wstring out;
		IHTMLElement* pParent = nullptr;
		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		assert(pRange);
		_bstr_t bWord(wTerm.c_str());
		VARIANT_BOOL vBool;
		wstring wt = chaj::DOM::HTMLTxtRange_get_text(pRange);
		pRange->findText(bWord, 1000000, 2, &vBool);
		_bstr_t bstrClassVal(wNewStat.c_str());
		_bstr_t bstrChar(L"character");
		long lVal;
		
		while (vBool == VARIANT_TRUE)
		{
			pRange->parentElement(&pParent);
			assert(pParent != nullptr);
			wstring wCurTerm = GetAttributeValue(pParent, L"lwtTerm");
			if (wTerm == wCurTerm)
				pParent->put_className(bstrClassVal.GetBSTR());

			pRange->move(bstrChar, 1, &lVal);
			pRange->findText(bWord, 1000000, 0, &vBool);
		}
	}
	void UpdatePageSpans3(IHTMLDocument2* pDoc, const wstring& wTerm, const wstring& wNewStatus)
	{
		if (!pTW) return;

		IDispatch *pRoot = NULL, *pCur = NULL;
		wstring wStatClass = L"lwtStat";
		wStatClass.append(wNewStatus);
		
		pTW->get_root(&pRoot);
		if (!pRoot) return;

		pTW->putref_currentNode(pRoot);
		pCur = pRoot;

		do
		{
			IHTMLElement* pElement = GetAlternateInterface<IDispatch,IHTMLElement>(pCur);
			if (!pElement) break;

			wstring wCurTerm = GetAttributeValue(pElement, L"lwtTerm");
			if (wCurTerm == wTerm)
			{
				_bstr_t bstrClassVal(wStatClass.c_str());
				pElement->put_className(bstrClassVal.GetBSTR());
			}

			pElement->Release();
			pTW->nextNode(&pCur);
		} while (pCur); 
	}
	static long FilterNodes_LWTTermAndSent(IDispatch* pNode)
	{
		IHTMLElement* pElement = GetAlternateInterface<IDispatch,IHTMLElement>(pNode);
		if (!pElement)
			return 3;

		wstring wTerm(GetAttributeValue(pElement, L"lwtTerm"));
		wchar_t* wczSent = L"lwtSent";

		if (wTerm.size() > 0)
		{
			pElement->Release();
			return 1;
		}
		else
		{
			_bstr_t bstrStat(L"");
			HRESULT hr = pElement->get_className(bstrStat.GetAddress());
			pElement->Release();
			if (bstrStat.length() > 0 && wcsicmp(bstrStat.GetBSTR(), wczSent) == 0)
				return FILTER_ACCEPT;
		}
		
		return FILTER_SKIP;
	}
	static long FilterNodes_Scripts(IDispatch* pNode)
	{
		IHTMLElement* pElement = GetAlternateInterface<IDispatch,IHTMLElement>(pNode);
		if (!pElement)
			return FILTER_SKIP;

		BSTR bstrTag;
		pElement->get_tagName(&bstrTag);
		wstring wstrTag(bstrTag);
		SysFreeString(bstrTag);
		if (wcsicmp(wstrTag.c_str(), L"script") == 0)
			return FILTER_ACCEPT;
		else
			return FILTER_SKIP;
	}
	static long FilterNodes_LWTTerm(IDispatch* pNode)
	{
		IHTMLElement* pElement = GetAlternateInterface<IDispatch,IHTMLElement>(pNode);
		if (!pElement)
			return 3;

		wstring wTerm(GetAttributeValue(pElement, L"lwtTerm"));
		pElement->Release();

		if (wTerm.size() > 0)
			return 1;
		else
			return 3;
	}
	void UpdatePageSpans(IHTMLDocument2* pDoc, const wstring& wTerm, const wstring& wNewStatus)
	{
		return UpdatePageSpans4(pDoc, wTerm, wNewStatus);
		IHTMLElementCollection* iEC = NULL;
		HRESULT hr = pDoc->get_all(&iEC);
		if (FAILED(hr))
		{
			mb("Could not update page elements.", "288fhuidih");
			return;
		}

		IDispatch* iDisp = NULL;
		VARIANT vTagName;
		vTagName.vt = VT_BSTR;
		_bstr_t bstrTagName("span");
		vTagName.bstrVal = bstrTagName.GetBSTR();
		hr = iEC->tags(vTagName, &iDisp);
		iEC->Release();
		if (FAILED(hr))
		{
			mb("Could not get list of span elements.", "302yfhjud");
			return;
		}

		IHTMLElementCollection* iECSpans;
		hr = iDisp->QueryInterface(IID_IHTMLElementCollection, (void**)&iECSpans);
		iDisp->Release();
		if (FAILED(hr))
		{
			mb("Could not get an element collection interface.", "311uol;jij");
			return;
		}

		long lLength = 0;
		hr = iECSpans->get_length(&lLength);
		if (FAILED(hr))
		{
			iECSpans->Release();
			mb("Could not get collection length.", "319dafesf");
			return;
		}

		if (lLength == 0)
		{
			iECSpans->Release();
			return;
		}

		VARIANT vIndex;
		vIndex.vt = VT_I4;
		IDispatch* pDispElem;
		IHTMLElement* pSpan;
		VARIANT vSubIndex;
		vSubIndex.vt = VT_I4;
		vSubIndex.lVal = 0;
		wstring wNewStat;
		wNewStat.append(L"lwtStat");
		wNewStat.append(wNewStatus);
		VARIANT varNewClassValue;
		varNewClassValue.vt = VT_BSTR;
		for (long i = 0; i < lLength; ++i)
		{
			vIndex.lVal = i;
			hr = iECSpans->item(vIndex, vSubIndex, (IDispatch**)&pDispElem);
			if (FAILED(hr))
			{
				continue;
			}
			hr = pDispElem->QueryInterface(IID_IHTMLElement, (void**)&pSpan);
			pDispElem->Release();
			if (FAILED(hr))
			{
				continue;
			}
			wstring wAtt(L"lwtTerm");
			wstring wAttValue = GetAttributeValue(pSpan, wAtt);
			if (wAttValue == wTerm)
			{
				//_bstr_t bstrClass("class");
				_bstr_t bstrClassVal(wNewStat.c_str());
				//varNewClassValue.bstrVal = bstrClassVal.GetBSTR();
				//hr = pSpan->setAttribute(bstrClass.GetBSTR(), varNewClassValue, 0);
				hr = pSpan->put_className(bstrClassVal.GetBSTR());//, varNewClassValue, 0);
			}
			pSpan->Release();
		}

		iECSpans->Release();
	}
	//wstring GetAttributeValue(IHTMLElement* pEl, const wstring& wAtt)
	//{
	//	_bstr_t bstrAtt(wAtt.c_str());
	//	VARIANT varAttValue;
	//	varAttValue.vt = VT_BSTR;
	//	HRESULT hr = pEl->getAttribute(bstrAtt.GetBSTR(), 2, &varAttValue);
	//	//{
	//	//	mb(L"Error upon attempting to retrieve attribute value.", wAtt);
	//	//	return L"";
	//	//}
	//	if (FAILED(hr))
	//		return wstring(L"");
	//	else
	//	{
	//		if (varAttValue.vt == VT_NULL)
	//			return wstring(L"");
	//		else
	//			return wstring(varAttValue.bstrVal);
	//	}
	//}
	void SendLwtCommand(const wstring& wCommand, IHTMLDocument2* pDoc)
	{
		IHTMLElement* pBhoCommand = chaj::DOM::GetElementFromId(L"lwtBhoCommand", pDoc);
		assert(pBhoCommand);
		chaj::DOM::SetAttributeValue(pBhoCommand, L"value", wCommand);
		pBhoCommand->click();
		pBhoCommand->Release();
	}
	void OnTableSetChange(IHTMLDocument2* pDoc, bool bReload = true)
	{
		GetLgID();
		OnLanguageChange(pDoc, bReload);
	}
	void OnLanguageChange(IHTMLDocument2* pDoc, bool bReload = true)
	{
		long lngLgID = wcstol(wLgID.c_str(), NULL, 10);

		if (lngLgID == 0 || lngLgID == LONG_MAX || lngLgID == LONG_MIN)
			return bReload ? ReloadWebpage(pDoc) : void();
		        
		wstring wQuery = L"select LgRegexpWordCharacters, LgRegexpSplitSentences, COALESCE(LgExceptionsSplitSentences, ''), LgSplitEachChar, COALESCE(LgDict1URI, ''), COALESCE(LgDict2URI, ''), COALESCE(LgGoogleTranslateURI, '') from ";
		wQuery += wTableSetPrefix;
		wQuery.append(L"languages where LgID = ");
		wQuery += wLgID;
		if (!DoExecuteDirect(_T("235ywiehf"), wQuery))
			return bReload ? ReloadWebpage(pDoc) : void();

		if (!tStmt.fetch_next())
			mb("Could not retrieve Word Characters. 249ydhfu.");
		else
		{
			WordChars = tStmt.field(1).as_string();
			SentDelims = tStmt.field(2).as_string();
			SentDelimExcepts = RegexEscape(tStmt.field(3).as_string());
			tStmt.field(4).as_string() == L"0" ? bWithSpaces = true : bWithSpaces = false;
			wstrDict1 = tStmt.field(5).as_string();
			wstrDict2 = tStmt.field(6).as_string();
			wstrGoogleTrans = tStmt.field(7).as_string();
			RegexEscape(SentDelimExcepts);
			auto it = cacheMap.find(wTableSetPrefix + wstring(L"~~sep~~") + wLgID);
			if (it != cacheMap.end())
				cache = it->second;
			else
			{
				cache = new LWTCache(wLgID, wTableSetPrefix);
				cacheMap.insert(pair<wstring,LWTCache*>(wTableSetPrefix + wstring(L"~~~sep~~~") + wLgID, cache));
			}

			if (bReload)
			{
				ReloadWebpage(pDoc);
			}
		}
		tStmt.free_results();
	}
	wstring RegexEscape(const wstring& in)
	{
		return regex_replace(in, wregex(L"([$.?\\^*+()])"), L"\\$1");
	}
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
	{
		if (riid==IID_IUnknown)
			*ppv=static_cast<LwtBho*>(this);
		else if (riid==IID_IObjectWithSite)
			*ppv=static_cast<IObjectWithSite*>(this); 
		else if (riid==IID_IDispatch)
			*ppv=static_cast<IDispatch*>(this);
		else
			return E_NOINTERFACE;
		















		AddRef();
		return S_OK;
	}
	ULONG STDMETHODCALLTYPE AddRef() {InterlockedIncrement(&gref); return InterlockedIncrement(&ref);}
	ULONG STDMETHODCALLTYPE Release() {int tmp=InterlockedDecrement(&ref); if (tmp==0) delete this; InterlockedDecrement(&gref); return tmp;}

	HRESULT STDMETHODCALLTYPE GetTypeInfoCount(unsigned int FAR* pctinfo) {*pctinfo=1; return NOERROR;}
	HRESULT STDMETHODCALLTYPE GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR*  ppTInfo)
	{
		return NOERROR;
	}
	HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId)
	{
		return NOERROR;
	}

	//wregex_iterator DoWRegexIter(wstring& pattern, wstring& src, wregex::flag_type flags = regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript)
	//{
	//	wregex wrgx(pattern, flags);
	//	wregex_iterator regit(src.begin(), src.end(), wrgx);
	//	return regit;
	//}
	//void DoWRegexIter2(wregex_iterator& it, wstring& pattern, wstring& src, wregex::flag_type flags = regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript)
	//{
	//	wregex wrgx(pattern, flags);
	//	it = wregex_iterator(src.begin(), src.end(), wrgx);
	//	//return regit;
	//}

	/*
		TagNotTag
		Input Condition: no isolated < characters that don't represent a valid html tag, such as in javascript code
	*/
	int FindProperCloseTagPos(const wstring& in, int nOpenTagPos)
	{
		assert(in.find(L"<", nOpenTagPos) == nOpenTagPos);
		wstring::size_type i = nOpenTagPos + 1;

		do
		{
			wchar_t wCur = in[i];
			if (wCur == '\"' || wCur == '\'')
			{
				wstring::size_type j = i + 1;
				do
				{
					wchar_t jCur = in[j];
					if (jCur == '\\')
						++j;
					else if (jCur == wCur)
					{
						i = j;
						break;
					}
					++j;
				} while (j < in.size());
			}
			if (wCur == '>')
				return i;

			++i;
		} while (i < in.size());

		return -1;
	}
	void TagNotTag(vector<wstring>& chunks, wstring& part)
	{
		int pos = part.find(L"<"), pos2 = 0;
		while (pos != wstring::npos)
		{
			if (pos != pos2) // some non tag data
				chunks.push_back(wstring(part, pos2, pos-pos2));

			if (part.find(L"<!--", pos) == pos)
			{
				pos2 = part.find(L"-->", pos);
				pos2 = part.find(L">", pos2);
			}
			else
				pos2 = FindProperCloseTagPos(part, pos);//part.find(L">", pos);

			if (pos2 != wstring::npos)
			{
				chunks.push_back(wstring(part, pos, pos2 - pos + 1));
				pos = part.find(L"<", ++pos2);
			}
			else // expected close tag not found
			{
				chunks.push_back(wstring(part, pos)); // take the rest of the part from false open tag on
				pos = pos2;
			}
		}
		if (pos2 != wstring::npos) // 
			chunks.push_back(wstring(part, pos2));
		//wregex_iterator regit(part.begin(), part.end(), TagNotTagWrgx);
		//while (regit != rend)
		//{
		//	chunks.push_back(regit->str());
		//	++regit;
		//}
	}
	bool TagIsSelfClosed(const wstring& wTag)
	{
		wstring::size_type pos = wTag.find_last_not_of(L"> ");
		if (wTag[pos] != '/')
			return false;
		if (wTag.find(L">", pos) == wstring::npos)
			return false;

		return true;
	}
	void SplitChunksByScript(vector<wstring>& chunks, wstring& wBodyPart)
	{
		if (wBodyPart.find(L"script") == wstring::npos)
			return TagNotTag(chunks, wBodyPart);

		wstring::size_type pos1 = 0, pos2 = 0;
		bool bOpenScriptContext = false; // have seen open <script> but not yet </script>

		wstring wrgxPtn1 = L"< */? *script.*?>";
		wregex wrgx1(wrgxPtn1, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regScripts(wBodyPart.begin(), wBodyPart.end(), wrgx1);

		if (regScripts == rend) // a script-tagless page
			return TagNotTag(chunks, wBodyPart);

		int nEnd;

		while (regScripts != rend)
		{
			wstring wScript = regScripts->str();
			if (bOpenScriptContext == true) // make sure we have a true close tag, and not an embedded script within javascript
				if (wScript.c_str()[wScript.find_first_of(L"/s")] == 's') // '/' did not prcede 's'
				{
					regScripts++;
					continue;
				}
			pos2 = wBodyPart.find(wScript, pos1);
			nEnd = FindProperCloseTagPos(wBodyPart, pos2);
			if (nEnd-pos2+1 != wScript.size())
				wScript = wstring(wBodyPart, pos2, nEnd-pos2+1);

			if (pos1 != pos2) // some content prior to the found script tag
			{
				if (bOpenScriptContext == true)
					chunks.push_back(wstring(wBodyPart, pos1, pos2 - pos1));
				else // this is an open script tag we found, what precedes is not javascript
					TagNotTag(chunks, wstring(wBodyPart, pos1, pos2 - pos1));
			}
			
			chunks.push_back(wScript);

			pos1 = pos2 + wScript.size();
			++regScripts;
			bOpenScriptContext = !bOpenScriptContext;
			if (bOpenScriptContext && TagIsSelfClosed(wScript))
				bOpenScriptContext = !bOpenScriptContext;
		}
		if (pos1 < wBodyPart.size())
			TagNotTag(chunks, wstring(wBodyPart, pos1, wBodyPart.size()-pos1));
	}
	/*
		PopulateChunks
		Exit Condition: no empty string chunks, but 0 chunks is fine
	*/
	void SplitChunksByIframe(vector<wstring>& chunks, wstring& wBodyPart)
	{
		if (wBodyPart.find(L"iframe") == wstring::npos)
			return SplitChunksByScript(chunks, wBodyPart);

		wstring::size_type pos1 = 0, pos2 = 0;
		bool bOpenScriptContext = false; // have seen open <script> but not yet </script>

		wstring wrgxPtn1 = L"< */? *iframe.*?>";
		wregex wrgx1(wrgxPtn1, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regScripts(wBodyPart.begin(), wBodyPart.end(), wrgx1);

		if (regScripts == rend) // a script-tagless page
			return SplitChunksByScript(chunks, wBodyPart);

		int nEnd;

		while (regScripts != rend)
		{
			wstring wScript = regScripts->str();
			if (bOpenScriptContext == true) // make sure we have a true close tag, and not an embedded script within javascript
				if (wScript.c_str()[wScript.find_first_of(L"/i")] == 's') // '/' did not prcede 's'
				{
					regScripts++;
					continue;
				}
			pos2 = wBodyPart.find(wScript, pos1);
			nEnd = FindProperCloseTagPos(wBodyPart, pos2);
			if (nEnd-pos2+1 != wScript.size())
				wScript = wstring(wBodyPart, pos2, nEnd-pos2+1);

			if (pos1 != pos2) // some content prior to the found script tag
			{
				if (bOpenScriptContext == true)
					chunks.push_back(wstring(wBodyPart, pos1, pos2 - pos1));
				else // this is an open script tag we found, what precedes is not javascript
					SplitChunksByScript(chunks, wstring(wBodyPart, pos1, pos2 - pos1));
			}
			
			chunks.push_back(wScript);

			pos1 = pos2 + wScript.size();
			++regScripts;
			bOpenScriptContext = !bOpenScriptContext;
			if (bOpenScriptContext && TagIsSelfClosed(wScript))
				bOpenScriptContext = !bOpenScriptContext;
		}
		if (pos1 < wBodyPart.size())
			SplitChunksByScript(chunks, wstring(wBodyPart, pos1, wBodyPart.size()-pos1));
	}

	void SplitChunksByComment(vector<wstring>& chunks, wstring& wBodyPart)
	{
		// this code does not support html that contains elements resembling comments like:
		// <input type="text" value="<!--this will in fact show to user correctly: not a comment-->" />

		int pos1, pos2;

		for (pos1 = 0, pos2 = wBodyPart.find(L"<!--", pos1); pos2 != wstring::npos; pos2 = wBodyPart.find(L"<!--", pos1))
		{
			if (pos2 != pos1)
				SplitChunksByIframe(chunks, wstring(wBodyPart, pos1, pos2 - pos1));

			pos1 = wBodyPart.find(L"-->", pos2) + 3;
			chunks.push_back(wstring(wBodyPart, pos2, pos1 - pos2));
		}

		SplitChunksByIframe(chunks, wstring(wBodyPart, pos1));
	}
	void PopulateChunks(vector<wstring>& chunks, wstring& wFullBody)
	{
		SplitChunksByComment(chunks, wFullBody);
	}

	/* WARNING */// the function IsChunkVisible is highly dependent a relationship among its lines of code
	// alter only with understanding
	bool IsChunkVisible(const vector<wstring>& chunks, int nChunkIndex)
	{
		if (chunks[nChunkIndex].c_str()[0] == L'<') // chuck definitely tag data
			return false;

		if (nChunkIndex == 0) // the following false case will not apply
			return true;

		if (TagIsSelfClosed(chunks[nChunkIndex-1]))
			return true;

		if (wcsncmp(chunks[nChunkIndex-1].c_str(), L"<script", 7) == 0 || \
			wcsncmp(chunks[nChunkIndex-1].c_str(), L"<style", 6) == 0 || \
			wcsncmp(chunks[nChunkIndex-1].c_str(), L"<iframe", 7) == 0 || \
			wcsncmp(chunks[nChunkIndex-1].c_str(), L"<noframe", 8) == 0 || \
			wcsncmp(chunks[nChunkIndex-1].c_str(), L"<textarea", 9) == 0)
			return false;
			
		return true;
	}
	bool ChunkIsPTag(vector<wstring>& chunks, int nChunkIndex)
	{
		wstring wrgxPtn = L"</?[ ]*p[ >/]";
		wregex wrgx(wrgxPtn, regex_constants::icase | regex_constants::ECMAScript);
		wregex_iterator regit(chunks[nChunkIndex].begin(), chunks[nChunkIndex].end(), wrgx);// = DoWRegexIter(wrgx, chunks[nChunkIndex]);
		wregex_iterator rend;
		if (regit != rend)
			return true;
		else
			return false;
	}

	wstring GetVisibleBodyWithTagBookmarks(vector<wstring>& chunks)
	{
		wstring wVisibleBody;
		bool bIsVisibleChunk; // must be properly set each loop iteration

		for (decltype(chunks.size()) i = 0; i < chunks.size(); ++i)
		{
			bIsVisibleChunk = IsChunkVisible(chunks, i);

			if (bIsVisibleChunk == true)
				wVisibleBody += chunks[i];
			else
			{
				if (ChunkIsPTag(chunks, i) == true)
					wVisibleBody += wPTagBookmark + to_wstring(i) + wPTagBookmark;
				else
					wVisibleBody += wHiddenChunkBookmark + to_wstring(i) + wHiddenChunkBookmark;
			}
		}

		return wVisibleBody;
	}

	void GetEntityMap(std::map<std::wstring,wchar_t>& entities)
	{
		entities[L"quot"]=(wchar_t)0x0022;
		entities[L"amp"]=(wchar_t)0x0026;
		entities[L"apos"]=(wchar_t)0x0027;
		entities[L"lt"]=(wchar_t)0x003C;
		entities[L"gt"]=(wchar_t)0x003E;
		entities[L"nbsp"]=(wchar_t)0x00A0;
		entities[L"iexcl"]=(wchar_t)0x00A1;
		entities[L"cent"]=(wchar_t)0x00A2;
		entities[L"pound"]=(wchar_t)0x00A3;
		entities[L"curren"]=(wchar_t)0x00A4;
		entities[L"yen"]=(wchar_t)0x00A5;
		entities[L"brvbar"]=(wchar_t)0x00A6;
		entities[L"sect"]=(wchar_t)0x00A7;
		entities[L"uml"]=(wchar_t)0x00A8;
		entities[L"copy"]=(wchar_t)0x00A9;
		entities[L"ordf"]=(wchar_t)0x00AA;
		entities[L"laquo"]=(wchar_t)0x00AB;
		entities[L"not"]=(wchar_t)0x00AC;
		entities[L"shy"]=(wchar_t)0x00AD;
		entities[L"reg"]=(wchar_t)0x00AE;
		entities[L"macr"]=(wchar_t)0x00AF;
		entities[L"deg"]=(wchar_t)0x00B0;
		entities[L"plusmn"]=(wchar_t)0x00B1;
		entities[L"sup2"]=(wchar_t)0x00B2;
		entities[L"sup3"]=(wchar_t)0x00B3;
		entities[L"acute"]=(wchar_t)0x00B4;
		entities[L"micro"]=(wchar_t)0x00B5;
		entities[L"para"]=(wchar_t)0x00B6;
		entities[L"middot"]=(wchar_t)0x00B7;
		entities[L"cedil"]=(wchar_t)0x00B8;
		entities[L"sup1"]=(wchar_t)0x00B9;
		entities[L"ordm"]=(wchar_t)0x00BA;
		entities[L"raquo"]=(wchar_t)0x00BB;
		entities[L"frac14"]=(wchar_t)0x00BC;
		entities[L"frac12"]=(wchar_t)0x00BD;
		entities[L"frac34"]=(wchar_t)0x00BE;
		entities[L"iquest"]=(wchar_t)0x00BF;
		entities[L"Agrave"]=(wchar_t)0x00C0;
		entities[L"Aacute"]=(wchar_t)0x00C1;
		entities[L"Acirc"]=(wchar_t)0x00C2;
		entities[L"Atilde"]=(wchar_t)0x00C3;
		entities[L"Auml"]=(wchar_t)0x00C4;
		entities[L"Aring"]=(wchar_t)0x00C5;
		entities[L"AElig"]=(wchar_t)0x00C6;
		entities[L"Ccedil"]=(wchar_t)0x00C7;
		entities[L"Egrave"]=(wchar_t)0x00C8;
		entities[L"Eacute"]=(wchar_t)0x00C9;
		entities[L"Ecirc"]=(wchar_t)0x00CA;
		entities[L"Euml"]=(wchar_t)0x00CB;
		entities[L"Igrave"]=(wchar_t)0x00CC;
		entities[L"Iacute"]=(wchar_t)0x00CD;
		entities[L"Icirc"]=(wchar_t)0x00CE;
		entities[L"Iuml"]=(wchar_t)0x00CF;
		entities[L"ETH"]=(wchar_t)0x00D0;
		entities[L"Ntilde"]=(wchar_t)0x00D1;
		entities[L"Ograve"]=(wchar_t)0x00D2;
		entities[L"Oacute"]=(wchar_t)0x00D3;
		entities[L"Ocirc"]=(wchar_t)0x00D4;
		entities[L"Otilde"]=(wchar_t)0x00D5;
		entities[L"Ouml"]=(wchar_t)0x00D6;
		entities[L"times"]=(wchar_t)0x00D7;
		entities[L"Oslash"]=(wchar_t)0x00D8;
		entities[L"Ugrave"]=(wchar_t)0x00D9;
		entities[L"Uacute"]=(wchar_t)0x00DA;
		entities[L"Ucirc"]=(wchar_t)0x00DB;
		entities[L"Uuml"]=(wchar_t)0x00DC;
		entities[L"Yacute"]=(wchar_t)0x00DD;
		entities[L"THORN"]=(wchar_t)0x00DE;
		entities[L"szlig"]=(wchar_t)0x00DF;
		entities[L"agrave"]=(wchar_t)0x00E0;
		entities[L"aacute"]=(wchar_t)0x00E1;
		entities[L"acirc"]=(wchar_t)0x00E2;
		entities[L"atilde"]=(wchar_t)0x00E3;
		entities[L"auml"]=(wchar_t)0x00E4;
		entities[L"aring"]=(wchar_t)0x00E5;
		entities[L"aelig"]=(wchar_t)0x00E6;
		entities[L"ccedil"]=(wchar_t)0x00E7;
		entities[L"egrave"]=(wchar_t)0x00E8;
		entities[L"eacute"]=(wchar_t)0x00E9;
		entities[L"ecirc"]=(wchar_t)0x00EA;
		entities[L"euml"]=(wchar_t)0x00EB;
		entities[L"igrave"]=(wchar_t)0x00EC;
		entities[L"iacute"]=(wchar_t)0x00ED;
		entities[L"icirc"]=(wchar_t)0x00EE;
		entities[L"iuml"]=(wchar_t)0x00EF;
		entities[L"eth"]=(wchar_t)0x00F0;
		entities[L"ntilde"]=(wchar_t)0x00F1;
		entities[L"ograve"]=(wchar_t)0x00F2;
		entities[L"oacute"]=(wchar_t)0x00F3;
		entities[L"ocirc"]=(wchar_t)0x00F4;
		entities[L"otilde"]=(wchar_t)0x00F5;
		entities[L"ouml"]=(wchar_t)0x00F6;
		entities[L"divide"]=(wchar_t)0x00F7;
		entities[L"oslash"]=(wchar_t)0x00F8;
		entities[L"ugrave"]=(wchar_t)0x00F9;
		entities[L"uacute"]=(wchar_t)0x00FA;
		entities[L"ucirc"]=(wchar_t)0x00FB;
		entities[L"uuml"]=(wchar_t)0x00FC;
		entities[L"yacute"]=(wchar_t)0x00FD;
		entities[L"thorn"]=(wchar_t)0x00FE;
		entities[L"yuml"]=(wchar_t)0x00FF;
		entities[L"OElig"]=(wchar_t)0x0152;
		entities[L"oelig"]=(wchar_t)0x0153;
		entities[L"Scaron"]=(wchar_t)0x0160;
		entities[L"scaron"]=(wchar_t)0x0161;
		entities[L"Yuml"]=(wchar_t)0x0178;
		entities[L"fnof"]=(wchar_t)0x0192;
		entities[L"circ"]=(wchar_t)0x02C6;
		entities[L"tilde"]=(wchar_t)0x02DC;
		entities[L"Alpha"]=(wchar_t)0x0391;
		entities[L"Beta"]=(wchar_t)0x0392;
		entities[L"Gamma"]=(wchar_t)0x0393;
		entities[L"Delta"]=(wchar_t)0x0394;
		entities[L"Epsilon"]=(wchar_t)0x0395;
		entities[L"Zeta"]=(wchar_t)0x0396;
		entities[L"Eta"]=(wchar_t)0x0397;
		entities[L"Theta"]=(wchar_t)0x0398;
		entities[L"Iota"]=(wchar_t)0x0399;
		entities[L"Kappa"]=(wchar_t)0x039A;
		entities[L"Lambda"]=(wchar_t)0x039B;
		entities[L"Mu"]=(wchar_t)0x039C;
		entities[L"Nu"]=(wchar_t)0x039D;
		entities[L"Xi"]=(wchar_t)0x039E;
		entities[L"Omicron"]=(wchar_t)0x039F;
		entities[L"Pi"]=(wchar_t)0x03A0;
		entities[L"Rho"]=(wchar_t)0x03A1;
		entities[L"Sigma"]=(wchar_t)0x03A3;
		entities[L"Tau"]=(wchar_t)0x03A4;
		entities[L"Upsilon"]=(wchar_t)0x03A5;
		entities[L"Phi"]=(wchar_t)0x03A6;
		entities[L"Chi"]=(wchar_t)0x03A7;
		entities[L"Psi"]=(wchar_t)0x03A8;
		entities[L"Omega"]=(wchar_t)0x03A9;
		entities[L"alpha"]=(wchar_t)0x03B1;
		entities[L"beta"]=(wchar_t)0x03B2;
		entities[L"gamma"]=(wchar_t)0x03B3;
		entities[L"delta"]=(wchar_t)0x03B4;
		entities[L"epsilon"]=(wchar_t)0x03B5;
		entities[L"zeta"]=(wchar_t)0x03B6;
		entities[L"eta"]=(wchar_t)0x03B7;
		entities[L"theta"]=(wchar_t)0x03B8;
		entities[L"iota"]=(wchar_t)0x03B9;
		entities[L"kappa"]=(wchar_t)0x03BA;
		entities[L"lambda"]=(wchar_t)0x03BB;
		entities[L"mu"]=(wchar_t)0x03BC;
		entities[L"nu"]=(wchar_t)0x03BD;
		entities[L"xi"]=(wchar_t)0x03BE;
		entities[L"omicron"]=(wchar_t)0x03BF;
		entities[L"pi"]=(wchar_t)0x03C0;
		entities[L"rho"]=(wchar_t)0x03C1;
		entities[L"sigmaf"]=(wchar_t)0x03C2;
		entities[L"sigma"]=(wchar_t)0x03C3;
		entities[L"tau"]=(wchar_t)0x03C4;
		entities[L"upsilon"]=(wchar_t)0x03C5;
		entities[L"phi"]=(wchar_t)0x03C6;
		entities[L"chi"]=(wchar_t)0x03C7;
		entities[L"psi"]=(wchar_t)0x03C8;
		entities[L"omega"]=(wchar_t)0x03C9;
		entities[L"thetasym"]=(wchar_t)0x03D1;
		entities[L"upsih"]=(wchar_t)0x03D2;
		entities[L"piv"]=(wchar_t)0x03D6;
		entities[L"ensp"]=(wchar_t)0x2002;
		entities[L"emsp"]=(wchar_t)0x2003;
		entities[L"thinsp"]=(wchar_t)0x2009;
		entities[L"zwnj"]=(wchar_t)0x200C;
		entities[L"zwj"]=(wchar_t)0x200D;
		entities[L"lrm"]=(wchar_t)0x200E;
		entities[L"rlm"]=(wchar_t)0x200F;
		entities[L"ndash"]=(wchar_t)0x2013;
		entities[L"mdash"]=(wchar_t)0x2014;
		entities[L"lsquo"]=(wchar_t)0x2018;
		entities[L"rsquo"]=(wchar_t)0x2019;
		entities[L"sbquo"]=(wchar_t)0x201A;
		entities[L"ldquo"]=(wchar_t)0x201C;
		entities[L"rdquo"]=(wchar_t)0x201D;
		entities[L"bdquo"]=(wchar_t)0x201E;
		entities[L"dagger"]=(wchar_t)0x2020;
		entities[L"Dagger"]=(wchar_t)0x2021;
		entities[L"bull"]=(wchar_t)0x2022;
		entities[L"hellip"]=(wchar_t)0x2026;
		entities[L"permil"]=(wchar_t)0x2030;
		entities[L"prime"]=(wchar_t)0x2032;
		entities[L"Prime"]=(wchar_t)0x2033;
		entities[L"lsaquo"]=(wchar_t)0x2039;
		entities[L"rsaquo"]=(wchar_t)0x203A;
		entities[L"oline"]=(wchar_t)0x203E;
		entities[L"frasl"]=(wchar_t)0x2044;
		entities[L"euro"]=(wchar_t)0x20AC;
		entities[L"image"]=(wchar_t)0x2111;
		entities[L"weierp"]=(wchar_t)0x2118;
		entities[L"real"]=(wchar_t)0x211C;
		entities[L"trade"]=(wchar_t)0x2122;
		entities[L"alefsym"]=(wchar_t)0x2135;
		entities[L"larr"]=(wchar_t)0x2190;
		entities[L"uarr"]=(wchar_t)0x2191;
		entities[L"rarr"]=(wchar_t)0x2192;
		entities[L"darr"]=(wchar_t)0x2193;
		entities[L"harr"]=(wchar_t)0x2194;
		entities[L"crarr"]=(wchar_t)0x21B5;
		entities[L"lArr"]=(wchar_t)0x21D0;
		entities[L"uArr"]=(wchar_t)0x21D1;
		entities[L"rArr"]=(wchar_t)0x21D2;
		entities[L"dArr"]=(wchar_t)0x21D3;
		entities[L"hArr"]=(wchar_t)0x21D4;
		entities[L"forall"]=(wchar_t)0x2200;
		entities[L"part"]=(wchar_t)0x2202;
		entities[L"exist"]=(wchar_t)0x2203;
		entities[L"empty"]=(wchar_t)0x2205;
		entities[L"nabla"]=(wchar_t)0x2207;
		entities[L"isin"]=(wchar_t)0x2208;
		entities[L"notin"]=(wchar_t)0x2209;
		entities[L"ni"]=(wchar_t)0x220B;
		entities[L"prod"]=(wchar_t)0x220F;
		entities[L"sum"]=(wchar_t)0x2211;
		entities[L"minus"]=(wchar_t)0x2212;
		entities[L"lowast"]=(wchar_t)0x2217;
		entities[L"radic"]=(wchar_t)0x221A;
		entities[L"prop"]=(wchar_t)0x221D;
		entities[L"infin"]=(wchar_t)0x221E;
		entities[L"ang"]=(wchar_t)0x2220;
		entities[L"and"]=(wchar_t)0x2227;
		entities[L"or"]=(wchar_t)0x2228;
		entities[L"cap"]=(wchar_t)0x2229;
		entities[L"cup"]=(wchar_t)0x222A;
		entities[L"int"]=(wchar_t)0x222B;
		entities[L"there4"]=(wchar_t)0x2234;
		entities[L"sim"]=(wchar_t)0x223C;
		entities[L"cong"]=(wchar_t)0x2245;
		entities[L"asymp"]=(wchar_t)0x2248;
		entities[L"ne"]=(wchar_t)0x2260;
		entities[L"equiv"]=(wchar_t)0x2261;
		entities[L"le"]=(wchar_t)0x2264;
		entities[L"ge"]=(wchar_t)0x2265;
		entities[L"sub"]=(wchar_t)0x2282;
		entities[L"sup"]=(wchar_t)0x2283;
		entities[L"nsub"]=(wchar_t)0x2284;
		entities[L"sube"]=(wchar_t)0x2286;
		entities[L"supe"]=(wchar_t)0x2287;
		entities[L"oplus"]=(wchar_t)0x2295;
		entities[L"otimes"]=(wchar_t)0x2297;
		entities[L"perp"]=(wchar_t)0x22A5;
		entities[L"sdot"]=(wchar_t)0x22C5;
		entities[L"lceil"]=(wchar_t)0x2308;
		entities[L"rceil"]=(wchar_t)0x2309;
		entities[L"lfloor"]=(wchar_t)0x230A;
		entities[L"rfloor"]=(wchar_t)0x230B;
		entities[L"lang"]=(wchar_t)0x2329;
		entities[L"rang"]=(wchar_t)0x232A;
		entities[L"loz"]=(wchar_t)0x25CA;
		entities[L"spades"]=(wchar_t)0x2660;
		entities[L"clubs"]=(wchar_t)0x2663;
		entities[L"hearts"]=(wchar_t)0x2665;
		entities[L"diams"]=(wchar_t)0x2666;


	}

	void ReplaceHTMLEntitiesWithChars(wstring& in)
	{
		if (in.find(L"&") == wstring::npos)
			return;

		wstring wrgxPtn = L"&(?:(?:[a-zA-Z]+)|(?:#[0-9]+));";
		wregex wrgx(wrgxPtn, regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wrgx);// = DoWRegexIter(wrgx, in);
		wregex_iterator rend;

		if (regit == rend)
			return;

		wstring out(in);
		std::map<std::wstring,wchar_t> entities;
		GetEntityMap(entities);

		while (regit != rend)
		{
			wstring found = regit->str();
			wstring entity(found, 1, found.size()-2);
			if (entities.count(entity) == 1)
			{
				if (wcsicmp(entity.c_str(), L"amp") == 0)
					//wstring_replaceAll(out, found, L"&#38;");
					wstring_replaceAll(out, found, L"~A~0~A~;");
				else if (wcsicmp(entity.c_str(), L"lt") == 0)
					//wstring_replaceAll(out, found, L"&#60;");
					wstring_replaceAll(out, found, L"~L~0~L~");
				else
					wstring_replaceAll(out, found, to_wstring(entities[entity]));
			}
			else
			{
				if (entity.c_str()[0] == '#')
				{
					wstring wNum(entity, 1, entity.size() - 1);
					UINT uNum = stoul(wNum);
					if (uNum == 60)
						wstring_replaceAll(out, found, L"~L~0~L~");
					else if (uNum == 38)
						wstring_replaceAll(out, found, L"~A~0~A~");
					else
					{
						wstring wstrChar(to_wstring((wchar_t)uNum));
						wstring_replaceAll(out, found, wstrChar);
					}
				}
				else
				{
					mb(L"Encountered unrecognized HTML Entity.", found);
				}
			}

			++regit;
		}
		
		in = out;
	}

	void UpdateCacheMWFragments(const wstring& wMWTerm)
	{
		if (bWithSpaces)
		{
			int pos = -1; //offset to jive with do loop

			do
			{
				++pos;
				pos = wMWTerm.find(L" ", pos);
				usetCacheMWFragments.insert(wstring(wMWTerm, 0, pos));
			} while (pos != wstring::npos);
		}
		else
		{
			for (wstring::size_type i = 1; i <= wMWTerm.size(); ++i)
				usetCacheMWFragments.insert(wstring(wMWTerm, 0, i));
		}
	}

	// GetUncacedSubset expects candidate words to be in lowercase already
	void GetUncachedSubset(vector<wstring>& vListUncached, const vector<wstring>& vCandidateList)
	{
		for (wstr_c_it itWord = vCandidateList.begin(), itWordEnd = vCandidateList.end(); itWord != itWordEnd; ++itWord)
		{
			unordered_map<wstring,TermRecord>::const_iterator it = cache->find(*itWord);
			if (it == cache->end())
			{
				//mb("not found");
				vListUncached.push_back(*itWord);
			}
			//else
			//{
			//	mb("the item");
			//	mb(*itWord);
			//	mb(to_wstring((*itWord).size()));
			//	UpdateCacheMWFragments(*itWord);
			//	//mb("found");
			//}
		}
	}

	void GetListOfWordsLC(vector<wstring>& vList, wstring& src)
	{
		wstring rgxPtn = L"[";
		rgxPtn += WordChars + L"]+";
		wregex wrgx(rgxPtn, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regit(src.begin(), src.end(), wrgx), rend;
		//for (wregex_iterator regit = DoWRegexIter(rgxPtn, src), rend; regit != rend; ++regit)
		while(regit != rend)
		{
			vList.push_back(chaj::str::wstring_tolower(regit->str()));
			++regit;
		}
	}
	void ParseSentencesToMWTerms(list<wstring>& lstItems, const TokenStruct& tokens)
	{
		for (vector<Token_Sentence>::size_type i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			for (vector<Token_Word>::size_type j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				wstring vTerm;
				int nSkipCount = 0;

				for (vector<Token_Word>::size_type k = j; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
				{
					if (tokens[i].isWordDigested(k) == false)
					{
						++nSkipCount;
						continue;
					}

					if (k > j && bWithSpaces == true)
						vTerm.append(L" ");
					vTerm.append(tokens[i][k]);

					if (k > j)
						lstItems.push_back(vTerm);

					//assert(cache->count(tokens[i][j]) == 1);
					//TermRecord rec = cache[tokens[i][j]];
					//if (rec.nTermsBeyond == 0)
					//	break;
				}
			}
		}
		lstItems.sort();
		lstItems.unique();
	}
	// the only if means only if the potential term might be in the LWT words table
	//void ParseSentencesToMWTermsOnlyIf(const TokenStruct& tknCanonical)
	//{
	//	for (int i = 0; i < ts.size(); ++i)
	//	{
	//		if (ts[i].bDigested == false)
	//			continue;

	//		for (int j = 0; j < ts[i].size(); ++j)
	//		{
	//			if (ts[i].isWordDigested(j) == false)
	//				continue;

	//			wstring vTerm;
	//			int nSkipCount = 0;

	//			for (int k = j; k < j + 9 + nSkipCount && k < ts[i].size(); ++k)
	//			{
	//				if (ts[i].isWordDigested(k) == false)
	//				{
	//					++nSkipCount;
	//					continue;
	//				}

	//				wstring wWord = WordSansBookmarks(ts[i][k]);

	//				if (k > j && bWithSpaces == true)
	//					vTerm.append(L" ");
	//				vTerm += wWord;

	//				if (k > j)
	//					vTerms.push_back(vTerm);

	//				assert(cache->count(wWord) == 1);
	//				TermRecord rec = cache[wWord];
	//				if (rec.nTermsBeyond == 0)
	//					break;
	//			}
	//		}
	//	}
	//}
	//// the only if means only if the potential term might be in the LWT words table
	//void ParseSentencesToMWTermsOnlyIf_SansBookmarks(vector<wstring> vTerms, const vector<vector<wstring>>& vSentences)
	//{
	//	for (int i = 0; i < vSentences.size(); ++i)
	//	{
	//		for (int j = 0; j < vSentences[i].size(); ++j)
	//		{
	//			wstring vTerm;
	//			for (int k = j; k < j + 9 && k < vSentences[i].size(); ++k)
	//			{
	//				if (k > j && bWithSpaces == true)
	//					vTerm.append(L" ");
	//				
	//				wstring wWord = WordSansBookmarks(vSentences[i][k]);

	//				vTerm += wWord;

	//				if (k > 0)
	//					vTerms.push_back(vTerm);

	//				TermRecord rec = cache[wWord];
	//				if (rec.nTermsBeyond == 0)
	//					break;
	//			}
	//		}
	//	}
	//}

	void GetNonDuplicatesSubset(vector<wstring>& vOut, const vector<wstring>& vIn)
	{
		list<wstring> lstNoDups(vIn.begin(), vIn.end());
		lstNoDups.sort();
		lstNoDups.unique();
		vOut = vector<wstring>(lstNoDups.begin(), lstNoDups.end());
	}
	void UpdateTerminationCache(const vector<wstring>& vWords)
	{
		vector<wstring> vWordsNoDups;
		GetNonDuplicatesSubset(vWordsNoDups, vWords);

		for (vector<wstring>::size_type i = 0; i < vWordsNoDups.size(); ++i)
		{
			unordered_map<wstring,TermRecord>::iterator it = cache->find(vWordsNoDups[i]);
			if (it == cache->end())
				continue;
			
			wstring wQuery = L"select exists(select 1 from ";
			wQuery += wTableSetPrefix;
			wQuery.append(L"words where WoTextLC like '%");
			wQuery += vWordsNoDups[i];
			if (bWithSpaces == true)
				wQuery.append(L" %'");
			else
				wQuery.append(L"_%'");
			wQuery.append(L" LIMIT 1)");
			if (!DoExecuteDirect(_T("677idjfjf"), wQuery))
				break;

			if (tStmt.fetch_next())
				it->second.nTermsBeyond = tStmt.field(1).as_long() > 0 ? 1 : 0;

			tStmt.free_results();
		}
	}
	void UpdateCaches(const TokenStruct& tknCanonical)
	{
		list<wstring> lstUncachedWords = GetUncachedTokens(tknCanonical);
		EnsureRecordEntryForEachWord(lstUncachedWords);
		//UpdateTerminationCache(vWords);
		EnsureRecordEntryForEachMWTerm(tknCanonical);
	}
	//void UpdateCaches(const vector<vector<wstring>>& vSentencesVWords)
	//{
	//	vector<wstring> vWords;
	//	SentenceWordsToWordsMinusBookmarks(vWords, vSentencesVWords);

	//	EnsureRecordEntryForEachWord(vWords);
	//	UpdateTerminationCache(vWords);
	//	EnsureRecordEntryForEachMWTerm(vSentencesVWords);
	//}
	void InitializeWregexes()
	{
		wstring wWordSansBookmarksWrgxPtn;
		wWordSansBookmarksWrgxPtn.append(L"~.~[0-9]+~.~");
		//wWordSansBookmarksWrgxPtn.append(L"(?:");
		//wWordSansBookmarksWrgxPtn.append(wHiddenChunkBookmark);
		//wWordSansBookmarksWrgxPtn.append(L"|");
		//wWordSansBookmarksWrgxPtn.append(wPTagBookmark);
		//wWordSansBookmarksWrgxPtn.append(L")[0-9]+(?:");
		//wWordSansBookmarksWrgxPtn.append(wHiddenChunkBookmark);
		//wWordSansBookmarksWrgxPtn.append(L"|");
		//wWordSansBookmarksWrgxPtn.append(wPTagBookmark);
		//wWordSansBookmarksWrgxPtn.append(L")");
		wWordSansBookmarksWrgx.assign(wWordSansBookmarksWrgxPtn, regex_constants::optimize | regex_constants::ECMAScript);

		TagNotTagWrgx.assign(L"<!--.*?-->|<.*?>|[^<]+", regex_constants::optimize | regex_constants::ECMAScript);
	}
	wstring WordSansBookmarks(const wstring& in)
	{
		wstring out(in);
		wregex_iterator regit(in.begin(), in.end(), wWordSansBookmarksWrgx);

		while (regit != rend)
		{
			wstring wMatch = regit->str();
			out.replace(out.find(wMatch), wMatch.size(), L"");
			++regit;
		}

		return out;
	}
	void SentenceWordsToWordsMinusBookmarks(vector<wstring>& words, TokenStruct& tokens)
	{
		for (vector<Token_Sentence>::size_type i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			for (vector<Token_Word>::size_type j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				words.push_back(WordSansBookmarks(tokens[i][j]));
			}
		}
	}
	wstring EscapeSQLQueryValue(const wstring& wQuery)
	{
		wstring out(wQuery);
		EscapeSQLQueryValue(out);
		return out;
	}
	wstring& EscapeSQLQueryValue(wstring& wQuery)
	{
		int pos = wQuery.find_first_of(L"'\\");
		if (pos == wstring::npos)
			return wQuery;

		int pos2;
		pos = wQuery.find('\'');
		while (pos != wstring::npos)
		{
			if ((pos > 0 && wQuery[pos-1] != L'\\') || pos == 0)
			{
				wQuery.replace(pos, 1, L"\\\'");
				pos2 = pos + 2;
			}
			else
				pos2 = pos + 1;

			pos = wQuery.find('\'', pos2);
		}

		// final backslashes must come in even pairs
		pos = wQuery.rfind('\\');
		if (pos == wQuery.length() - 1) // ends with a backslash
		{
			if (pos == 1)
				wQuery = L"\\\\";
			else
			{
				pos2 = wQuery.find_last_not_of('\\');
				if ((pos - pos2) % 2 != 0)
				{
					wQuery.replace(pos, 1, L"\\\\");
				}
			}
		}

		return wQuery;
	}
	void CacheDBHitsWithListRemove(list<wstring>& lstItems, int nAtATime = 1000, bool bIsMultiWord = false)
	{
		int MAX_FIELD_WIDTH = 255;
		bool bWithList = false;
		vector<wstring> vQueries;
		list<wstring> lstCopy(lstItems);

		list<wstring>::iterator lend = lstCopy.end();
		list<wstring>::iterator iter = lstCopy.begin();
		for (list<wstring>::size_type i = 0; i < lstCopy.size(); i += nAtATime)
		{
			wstring wInList;

			for (list<wstring>::size_type j = i; j < lstCopy.size() && j < i + nAtATime; ++j)
			{
				wInList.append(L"'");
				wInList.append(EscapeSQLQueryValue(*iter));
				wInList.append(L"',");
				++iter;
			}

			wInList.append(L"'~~z~x~~Q_a_dummy_value'");
			wstring wQuery;
			wQuery.reserve(nAtATime * MAX_FIELD_WIDTH);
			wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from ");
			wQuery.append(wTableSetPrefix);
			wQuery.append(L"words where WoLgID = ");
			wQuery.append(wLgID);
			wQuery.append(L" AND WoTextLC in (");
			wQuery.append(wInList);
			wQuery.append(L")");
			vQueries.push_back(wQuery);
		}

		TRACE(L"%i queries created\n", vQueries.size());
		for (vector<wstring>::size_type i = 0; i < vQueries.size(); ++i)
		{
			EnterCriticalSection(&CS_UseDBConn);
			DoExecuteDirect(_T("1612isjdlfij"), vQueries[i]);
			while (tStmt.fetch_next())
			{
				wstring wLC = tStmt.field(1).as_string();
				TermRecord rec(tStmt.field(2).as_string());
				rec.wRomanization = tStmt.field(3).as_string();
				rec.wTranslation = tStmt.field(4).as_string();
				cache->insert(unordered_map<wstring,TermRecord>::value_type(wLC, rec));
				if (bIsMultiWord == true)
					UpdateCacheMWFragments(wLC);
				else
					lstItems.remove(wLC);
			}
			tStmt.free_results();
			LeaveCriticalSection(&CS_UseDBConn);
		}
	}
	list<wstring> GetUncachedTokens(const TokenStruct& tokens)
	{
		list<wstring> out;
		for (vector<Token_Sentence>::size_type i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested)
			{
				for (vector<Token_Word>::size_type j = 0; j < tokens[i].size(); ++j)
				{
					if (tokens[i].isWordDigested(j))
					{
						unordered_map<wstring,TermRecord>::const_iterator it = cache->find(tokens[i][j]);
						if (it == cache->end())
							out.push_back(tokens[i][j]);
					}
				}
			}
		}
		out.sort();
		out.unique();
		return out;
	}
	void EnsureRecordEntryForEachWord(list<wstring>& cWords)
	{
		const int MAX_MYSQL_INLIST_LEN = 2000;
		CacheDBHitsWithListRemove(cWords, MAX_MYSQL_INLIST_LEN);

		// cache misses with status 0
		for(wstring wTerm : cWords)
			cache->insert(unordered_map<wstring,TermRecord>::value_type(wTerm, TermRecord(L"0")));
	}
	void test5(vector<wstring>& v)
	{
		for (vector<wstring>::size_type i = 0; i < v.size(); ++i)
			mb(v[i]);
	}
	void EnsureRecordEntryForEachMWTerm(const TokenStruct& tknCanonical)
	{
		list<wstring> lstTerms;
		ParseSentencesToMWTerms(lstTerms, tknCanonical);

		//vector<wstring> vUncachedMWTerms;
		//GetUncachedSubset(vUncachedMWTerms, vTerms);

		const int nMaxMWTermsPerInList = 500;
		CacheDBHitsWithListRemove(lstTerms, nMaxMWTermsPerInList, true);
	}

	/*
		GetBodyContent
		Input conditions: pBody is NULL, pDoc is not NULL
		Exit condition: pDoc pointer equal to entrance value and has not had Release called
		If successful, pBody non-NULL
	*/
	void GetBodyContent(wstring& wBody, IHTMLElement*& pBody, IHTMLDocument2* pDoc)
	{
		wBody = L"";
		if (pBody != NULL || pDoc == NULL)
			return;
		
		IConnectionPointContainer* aCPC;
		HRESULT hr = pDoc->QueryInterface(IID_IConnectionPointContainer, (void**)&aCPC);
		if (FAILED(hr))
		{
			mb(_T("FindDocCPC fail"), _T("260suhdhf"));
			return;
		}

		IConnectionPoint* pHDEv2;
		hr = aCPC->FindConnectionPoint(DIID_HTMLDocumentEvents, &pHDEv2);
		aCPC->Release();
		if (FAILED(hr))
		{
			mb(_T("Could not find a place to register a click listener"), _T("267sezhdu -- FindHDEv fail"));
			return;
		}

		if (pHDEv != pHDEv2)
		{
			if (pHDEv != NULL)
			{
				pHDEv->Unadvise(Cookie_pHDev);
				pHDEv->Release();
			}
		
			pHDEv = pHDEv2;
			hr = pHDEv->Advise(reinterpret_cast<IDispatch*>(this), &Cookie_pHDev);
		}
		else
			pHDEv2->Release();
		if (FAILED(hr))
			mb("Error. Will not perceive mouse clicks on this page. Will try to annotate anyway.", "1000duwksm");

		hr = pDoc->get_body(&pBody);
		if (!pBody || FAILED(hr))
		{
			return;
		}
		
		//get body content, exluding <body> tags
		_bstr_t bstrBodyContent;
		hr = pBody->get_innerHTML(bstrBodyContent.GetAddress());
		if (FAILED(hr))
		{
			mb(L"get_innerHTML gave an error", L"137werts");
			return;
		}

		if (bstrBodyContent.length())
			wBody = to_wstring(bstrBodyContent.GetBSTR());

	}
	void ReceiveEvents(IHTMLDocument2* pDoc)
	{
		if (pDoc == nullptr)
			return;
		
		IConnectionPointContainer* aCPC;
		HRESULT hr = pDoc->QueryInterface(IID_IConnectionPointContainer, (void**)&aCPC);
		if (FAILED(hr))
		{
			mb(_T("FindDocCPC fail"), _T("260suhdhf"));
			return;
		}

		IConnectionPoint* pHDEv2;
		hr = aCPC->FindConnectionPoint(DIID_HTMLDocumentEvents, &pHDEv2);
		aCPC->Release();
		if (FAILED(hr))
		{
			mb(_T("Could not find a place to register a click listener"), _T("267sezhdu -- FindHDEv fail"));
			return;
		}

		if (pHDEv != pHDEv2)
		{
			if (pHDEv != NULL)
			{
				pHDEv->Unadvise(Cookie_pHDev);
				pHDEv->Release();
			}
		
			pHDEv = pHDEv2;
			hr = pHDEv->Advise(reinterpret_cast<IDispatch*>(this), &Cookie_pHDev);
		}
		else
			pHDEv2->Release();
		if (FAILED(hr))
			mb("Error. Will not perceive mouse clicks on this page. Will try to annotate anyway.", "1000duwksm");
	}
	void PSTS_BookmarkLevel(TokenStruct& tokens, const wstring& in)
	{
		wstring::size_type pos1 = 0, pos2 = 0;
		Token_Sentence ts(true);
		
		wstring regPtn;
		regPtn.append(L"(?:~[-ALP]~");
		regPtn.append(L"[0-9]+");
		regPtn.append(L"~[-ALP]~)+");

		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wrgx);
		wregex_iterator rend;

		while (regit != rend)
		{
			wstring wBookmarks = regit->str();
			assert(pos1 != wstring::npos && pos1 >= 0 && pos1 < in.size());
			pos2 = in.find(wBookmarks, pos1);
			assert(pos2 != wstring::npos); // regex pattern just found should be found again
			if (pos2 != pos1) // non-bookmark data found in sentence, try to parse as words
				PSTS_WordLevel2(ts, wstring(in, pos1, pos2-pos1));
			pos1 = pos2 + wBookmarks.size();

			ts.push_back(wBookmarks, false);

			++regit;
		}
		if (pos1 == 0) // no bookmarks present
		{
			PSTS_WordLevel2(ts, in);
			tokens.push_back(ts);
		}
		else
		{
			if (pos1 < in.size()) // we have data left after final bookmark
				PSTS_WordLevel2(ts, wstring(in, pos1, in.size()-pos1));

			tokens.push_back(ts);
		}

	}
	void PSTS_WordLevel2(Token_Sentence& ts, const wstring& in)
	{
		wstring::size_type pos1 = 0, pos2 = 0;

		// pattern: (?:(?:~-~#~-~)*[aWordChar])(?:~-~#~-~|[aWordChar])*
		// regex pattern regPtn ensures that it returns a true word (contains at least one lwt defined word character)
		// regPtns also ensures that if ~ is considered a valid wordChar, we won't end up at - and choke on a bookmark tag,
		// i.e. it eats whole bookmarks first at each comparison point, and it why the second appearance of [aWordChar]
		// is not [aWordChar]+
		wstring regPtn;
		regPtn.append(L"[");
		regPtn += WordChars;
		regPtn.append(L"]");
		if (bWithSpaces)
		{
			regPtn.append(L"+");
		}

		wregex wrgx(regPtn, regex_constants::icase | regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wrgx);
		wregex_iterator rend;

		while (regit != rend)
		{
			wstring wWord = regit->str();
			assert(pos1 != wstring::npos && pos1 >= 0 && pos1 < in.size());
			pos2 = in.find(wWord, pos1);
			assert(pos2 != wstring::npos); // regex pattern just found should be found again
			if (pos2 != pos1) // non-word data found in sentence
				ts.push_back(wstring(in, pos1, pos2-pos1), false);
			pos1 = pos2 + wWord.size();
			
			if (bWithSpaces)
				ts.push_back(wWord);
			else
			{
				for (wstring::size_type i = 0; i < wWord.size(); ++i)
					ts.push_back(wstring(wWord, i, 1));
			}

			++regit;
		}
		if (pos1 == 0) // only non-digest data in this sentence
		{
			ts.push_back(in, false);
		}
		else
		{
			if (pos1 < in.size()) // we have data left in the sentence we received that contains no words
				ts.push_back(wstring(in, pos1, in.size()-pos1), false);
		}

	}
	void PSTS_WordLevel(TokenStruct& tokens, const wstring& in)
	{
		wstring::size_type pos1 = 0, pos2 = 0;
		Token_Sentence ts(true);

		// pattern: (?:(?:~-~#~-~)*[aWordChar])(?:~-~#~-~|[aWordChar])*
		// regex pattern regPtn ensures that it returns a true word (contains at least one lwt defined word character)
		// regPtns also ensures that if ~ is considered a valid wordChar, we won't end up at - and choke on a bookmark tag,
		// i.e. it eats whole bookmarks first at each comparison point, and it why the second appearance of [aWordChar]
		// is not [aWordChar]+
		wstring regPtn;
		if (bWithSpaces)
		{
			regPtn.append(L"(?:(?:");
			regPtn += wHiddenChunkBookmark;
			regPtn.append(L"[0-9]+");
			regPtn += wHiddenChunkBookmark;
			regPtn.append(L")*[");
			regPtn += WordChars;
			regPtn.append(L"]?)(?:");
			regPtn += wHiddenChunkBookmark;
			regPtn.append(L"[0-9]+");
			regPtn += wHiddenChunkBookmark;
			regPtn.append(L"|[");
			regPtn += WordChars;
			regPtn.append(L"])*");
		}
		else
		{
			regPtn.append(L"[");
			regPtn += WordChars;
			regPtn.append(L"]");
		}

		//wstring sentLC = wstring_tolower(in);	
		wregex wrgx(regPtn, regex_constants::icase | regex_constants::ECMAScript);
		//wregex_iterator regit(sentLC.begin(), sentLC.end(), wrgx);
		wregex_iterator regit(in.begin(), in.end(), wrgx);
		wregex_iterator rend;
		int i = tokens.size();
		int ijk = 0;

		while (regit != rend)
		{
			wstring wWord = regit->str();
			assert(pos1 != wstring::npos && pos1 >= 0 && pos1 < in.size());
			pos2 = in.find(wWord, pos1);
			assert(pos2 != wstring::npos); // regex pattern just found should be found again
			if (pos2 != pos1) // non-word data found in sentence
				ts.push_back(wstring(in, pos1, pos2-pos1), false);
			pos1 = pos2 + wWord.size();
			
			if (bWithSpaces)
				ts.push_back(wWord);
			else
			{
				for (wstring::size_type i = 0; i < wWord.size(); ++i)
					ts.push_back(wstring(wWord, i, 1));
			}

			++regit;
		}
		if (pos1 == 0) // only non-digest data in this sentence
		{
			ts.push_back(in, false);
			tokens.push_back(ts);
		}
		else
		{
			if (pos1 < in.size()) // we have data left in the sentence we received that contains no words
				ts.push_back(wstring(in, pos1, in.size()-pos1), false);

			tokens.push_back(ts);
		}

	}
	void PSTS_LwtSentenceLevel(TokenStruct& tokens, const wstring& in)
	{
		wstring::size_type pos1 = 0, pos2 = 0;

		wstring rgxPtn = L"[^";
		rgxPtn += SentDelims + L"]+";
		wregex wrgx(rgxPtn);
		wregex_iterator regit(in.begin(), in.end(), wrgx);
		wregex_iterator rend;
		while (regit != rend)
		{
			wstring wSentence = regit->str();

			assert(pos1 != wstring::npos && pos1 >= 0 && pos1 < in.size()); 
			// pos1 != wstring::npos because that implies something found in the wregex wasn't again found in the src it was applied to
			// pos1 >= 0 because we set it to 0 and then only ever add, so just logially, if it's < 0, we didn't intend for that
			// pos1 < src.size() because the wregex_iterator should fair the while loop if we were ready to look beyond src length
			pos2 = in.find(wSentence, pos1);
			assert(pos2 != wstring::npos);
			// regex pattern just found should be found again
			if (pos2 != pos1) // there is a non-digested sentence-level token present
				tokens.push_back(wstring(in, pos1, pos2 - pos1));
			pos1 = pos2 + wSentence.size();

			//PSTS_WordLevel(tokens, wSentence);
			PSTS_BookmarkLevel(tokens, wSentence);

			++regit;
		}
		if (pos1 == 0) // no parseable sentences in src, all non-digested and passed through
		{
			tokens.push_back(in);
		}
		else
		{
			if (pos1 < in.size())
				tokens.push_back(wstring(in, pos1, in.size()-pos1));
		}
	}
	void PSTS_PTagLevel(TokenStruct& tokens, const wstring& in)
	{
		wstring wPat = wPTagBookmark + L"[0-9]+";
		wPat += wPTagBookmark;
		wregex wrgx(wPat);
		wregex_iterator regit(in.begin(), in.end(), wrgx), rend;
		wstring::size_type pos1 = 0, pos2 = 0;
		if (regit == rend)
			return PSTS_LwtSentenceLevel(tokens, in);
		while (regit != rend)
		{
			wstring wPTag = regit->str();
			pos2 = in.find(wPTag, pos1);
			if (pos1 != pos2)
				PSTS_LwtSentenceLevel(tokens, wstring(in, pos1, pos2-pos1));
			tokens.push_back(wPTag); // just push <P> tag onto token list at sentence level as not digested data
			pos1 = pos2 + wPTag.size();
			++regit;
		}
		if (pos1 < in.size())
		{
			PSTS_LwtSentenceLevel(tokens, wstring(in, pos1, in.size()-pos1));
		}
	}
	void ParseIntoSentencesAsTokenStruct(TokenStruct& tokens, wstring& src)
	{
		TRACE(L"%s", L"Calling ParseIntoSentencesAsTokenStruct\n");
		PSTS_PTagLevel(tokens, src);
		assert(tokens.vSentences.size() > 0);
		TRACE(L"%s", L"Leaving ParseIntoSentencesAsTokenStruct\n");
	}
	void ParseIntoSentencesAsWordList(vector<vector<wstring>> vSentencesVWords, wstring& src)
	{
		// pattern: (acceptable_sub_sequence|(?!~P~#~P~)[^a_sentence_separactor_char])+
		wstring rgxPtn = L"(";
		rgxPtn += SentDelimExcepts + L"|(?!";
		rgxPtn += wPTagBookmark + L"[0-9]+";
		rgxPtn += wPTagBookmark + L")[^";
		rgxPtn += SentDelims + L"])+";
		wregex wrgx(rgxPtn, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regit(src.begin(), src.end(), wrgx), rend;
		//for (wregex_iterator regit = DoWRegexIter(rgxPtn, src), rend; regit != rend; ++regit)
		while (regit != rend)
		{
			wstring wSentence = regit->str();
			vector<wstring> vWords;
			// pattern: ((~-~#~-~)*[aWordChar])(~-~#~-~|[aWordChar])*
			// regex pattern regPtn2 ensures that it returns a true word (contains at least one lwt defined word character)
			// regPtns also ensures that if ~ is considered a valid wordChar, we won't end up at - and choke on a bookmark tag,
			// i.e. it eats whole bookmarks first at each comparison point, and it why the second appearance of [aWordChar]
			// is not [aWordChar]+
			wstring regPtn2 = L"((";
			regPtn2 += wHiddenChunkBookmark;
			regPtn2.append(L"[0-9]+");
			regPtn2 += wHiddenChunkBookmark;
			regPtn2.append(L")*[");
			regPtn2 += WordChars;
			regPtn2.append(L"])(");
			regPtn2 += wHiddenChunkBookmark;
			regPtn2.append(L"[0-9]+");
			regPtn2 += wHiddenChunkBookmark;
			regPtn2.append(L"|[");
			regPtn2 += WordChars;
			regPtn2.append(L"])*");
			wregex wrgx(regPtn2, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
			wstring wSLC = chaj::str::wstring_tolower(wSentence);
			wregex_iterator regit2(wSLC.begin(), wSLC.end(), wrgx);
			//for (wregex_iterator regit2 = DoWRegexIter(regPtn2, wstring_tolower(wSentence)); regit2 != rend; ++regit2)
			while(regit2 != rend)
			{
				vWords.push_back(regit2->str());
				++regit2;
			}
			vSentencesVWords.push_back(vWords);

			++regit;
		}
	}
	void RestorePTagBookmarks(wstring& in, const vector<wstring>& chunks)
	{
		wstring out(in);
		int pos = 0;

		wstring wrgx = wPTagBookmark + L"[0-9]+";
		wrgx += wPTagBookmark;
		wregex wr(wrgx, regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wr);// = DoWRegexIter(wrgx, in);
		wregex_iterator rend;

		while (regit != rend)
		{
			wstring wBookmark = regit->str();
			wstring wChunkNum(wBookmark, nHiddenChunkOpenBookmarkLen, wBookmark.size()-(2*nHiddenChunkOpenBookmarkLen));
			int nChunkNum = 0;
			try
			{
				nChunkNum = stoi(wChunkNum);
			}
			catch (const std::invalid_argument&)
			{
				++regit;
				continue;
			}
			
			pos = out.find(wBookmark, pos);
			assert(pos != wstring::npos);
			out.replace(pos, wBookmark.size(), chunks[nChunkNum]);
			++pos;
			++regit;
		}

		in = out;
	}
	void RestoreAllBookmarks(wstring& in, const vector<wstring>& chunks)
	{
		wstring out(in);
		int pos = 0;

		RestorePTagBookmarks(out, chunks);

		wstring wrgx = wHiddenChunkBookmark + L"[0-9]+";
		wrgx += wHiddenChunkBookmark;
		wregex wr(wrgx, regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wr);// = DoWRegexIter(wrgx, in);
		wregex_iterator rend;

		while (regit != rend)
		{
			wstring wBookmark = regit->str();
			wstring wChunkNum(wBookmark, nHiddenChunkOpenBookmarkLen, wBookmark.size()-(2*nHiddenChunkOpenBookmarkLen));
			int nChunkNum = 0;
			try
			{
				nChunkNum = stoi(wChunkNum);
			}
			catch (const std::invalid_argument&)
			{
				++regit;
				continue;
			}
			
			pos = out.find(wBookmark, pos);
			assert(pos != wstring::npos);
			out.replace(pos, wBookmark.size(), chunks[nChunkNum]);
			++pos;
			++regit;
		}

		wstring_replaceAll(out, L"~A~0~A~", L"&amp;");
		wstring_replaceAll(out, L"~L~0~L~", L"&lt;");
		in = out;
	}
	void AppendWordSpan(wstring& out, const wstring& wStatus, const wstring& wWordOrigFormat, const wstring& wWordNoBookmarks, const wstring& wWordCanonical)
	{
		int loc = wWordOrigFormat.find(wWordNoBookmarks);
		if (loc == wstring::npos)
			out.append(wWordOrigFormat);
		else
		{
			if (loc > 0)
				out.append(wstring(wWordOrigFormat, 0, loc));
			out.append(L"<span class=\"lwtStat");
			out += wStatus;
			out.append(L"\" lwtTerm=\"");
			out.append(wWordCanonical);
			out.append(L"\">");
			out.append(wWordNoBookmarks);
			out.append(L"</span>");
			if (wWordOrigFormat.size() > loc + wWordNoBookmarks.size())
					out.append(wstring(wWordOrigFormat, loc + wWordNoBookmarks.size()));
		}
	}
	void AppendSoloWordSpan(wstring& out, const wstring& wAsAppended, const wstring& wWordCanonical, const TermRecord* pRec)
	{
		AppendWithSegues(out, wAsAppended, wWordCanonical, pRec);
	}
	void AppendMWSpan(wstring& out, const wstring& wTermCanonical, const TermRecord* pRec, const wstring& wWordCount)
	{
		out.append(L"<sup>");
		AppendWithSegues(out, wWordCount, wTermCanonical, pRec);
		out.append(L" </sup>");
	}
	void AppendWithSegues(wstring& out, const wstring& wAsAppended, const wstring& wTermCanonical, const TermRecord* pRec)
	{
		AppendTermIntro(out, wTermCanonical, pRec);
		out.append(wAsAppended);
		AppendTermAltro(out);
	}
	void AppendTermIntro(wstring& out, const wstring& wTermCanonical, const TermRecord* pRec)
	{
		out.append(L"<span class=\"lwtStat");
		out += pRec->wStatus;
		out.append(L"\" style=\"display:inline;margin:0px;color:black;\" lwtTerm=\"");
		out.append(wTermCanonical);
		out.append(L"\"");

		out.append(L" onmouseover=\"lwtmover('lwt");
		out.append(to_wstring(pRec->uIdent));
		out.append(L"', event, this);\" onmouseout=\"lwtmout(event);\"");

		out.append(L">");
	}
	void AppendTermAltro(wstring& out)
	{
		out.append(L"</span>");
	}
	void AppendJavascript(IHTMLDocument2* pDoc)
	{
		if (!pDoc)
		{
			TRACE(L"Argment error in AppendJavascript. 2513ishdllijj\n");
			return;
		}

		IHTMLElement* pHead = chaj::DOM::GetHeadFromDoc(pDoc); SmartCOMRelease scHead(pHead);
		IHTMLElement* pBody = chaj::DOM::GetBodyFromDoc(pDoc); SmartCOMRelease scBody(pBody);
		IHTMLElement* pScript = CreateElement(pDoc, L"script"); SmartCOMRelease scScript(pScript);
		IHTMLDOMNode* pScriptNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pScript); SmartCOMRelease scScriptNode(pScriptNode);
		if (!pHead || !pBody || !pScript || !pScriptNode)
		{
			TRACE(L"Unable to append lwt javascript properly.\n");
			return;
		}

		chaj::DOM::SetAttributeValue(pScript, L"type", L"text/javascript");
		chaj::DOM::SetElementInnerText(pScript, wJavascript);

		bool DoHead = false;
		if (DoHead)
		{
			IHTMLDOMNode* pHeadNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pHead); SmartCOMRelease scHeadNode(pHeadNode);
			IHTMLDOMNode* pNewScriptNode;
			HRESULT hr = pHeadNode->appendChild(pScriptNode, &pNewScriptNode);
			if (FAILED(hr))
				TRACE(L"Unable to append javacript child. 23328ijf\n");
			pNewScriptNode->Release();
		}
		else
		{
			IHTMLDOMNode* pBodyNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pBody); SmartCOMRelease scBodyNode(pBodyNode);
			IHTMLDOMNode* pNewScriptNode;
			HRESULT hr = pBodyNode->appendChild(pScriptNode, &pNewScriptNode);
			if (FAILED(hr))
				TRACE(L"Unable to append javacript child. 2546suhfuugs\n");
			pNewScriptNode->Release();
		}
	}
	/*
	AppendAnnotatedContent_MaintainBookmarks
	Input Condition: all words in vSentencesVWords are in the cache container
	*/
	void AppendAnnotatedContent_MaintainBookmarks
	(wstring& out, const TokenStruct& tokens, const TokenStruct& tknWithout, const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling AppendAnnotatedContent_MaintainBookmarks\n");
		for (vector<Token_Sentence>::size_type i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
			{
				out.append(tokens[i][0]);
				continue;
			}

			out.append(L"<span class=\"lwtSent\"></span>");

			for (vector<Token_Word>::size_type j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
				{
					out.append(tokens[i][j]);
					continue;
				}

				unordered_map<wstring,TermRecord>::const_iterator it = cache->find(tknCanonical[i][j]);
				assert(it != cache->end()); // there should always be a result
			
				//prepend any multiword term spans that start at this word
				if (it->second.nTermsBeyond != 0) //the term participates in multiwords
				{
					struct mwVals
					{
						mwVals(wstring wStat, wstring wWC, wstring wTerm) : wStatus(wStat), wWordCount(wWC), wMWTerm(wTerm) {};
						wstring wStatus;
						wstring wWordCount;
						wstring wMWTerm;
					};

					wstring curTerm = tknCanonical[i][j];
					if (usetCacheMWFragments.count(curTerm) != 0) //this terms starts some MWTerm
					{

						stack<mwVals> stkMWTerms;
						unsigned int nSkipCount = 0;
						// gather a stack of term statuses for tracked multiword terms that start at this word
						for (vector<Token_Word>::size_type k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
						{
							if (tokens[i].isWordDigested(k) == false)
							{
								++nSkipCount;
								continue;
							}

							if (bWithSpaces == true)
								curTerm.append(L" ");
							curTerm += tknCanonical[i][k];

							if (usetCacheMWFragments.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
								break;

							unordered_map<wstring,TermRecord>::const_iterator it = cache->find(curTerm);
							if (it != cache->end())
								stkMWTerms.push(mwVals(it->second.wStatus, to_wstring(k-j+1-nSkipCount), curTerm));
						}

						// append span tags for multiword terms, from longest to short
						while(stkMWTerms.size() > 0)
						{
							mwVals mwVCur = stkMWTerms.top();
							stkMWTerms.pop();
							out.append(L"<sup><span class=\"lwtStat");
							out += mwVCur.wStatus;
							out.append(L"\" lwtTerm=\"");
							out.append(mwVCur.wMWTerm);
							out.append(L"\">");
							out += mwVCur.wWordCount;
							out.append(L"</span></sup>");
						}
					}
				}

				AppendWordSpan(out, it->second.wStatus, tokens[i][j], tknWithout[i][j], tknCanonical[i][j]);
			}
		}
		TRACE(L"%s", L"Leaving AppendAnnotatedContent_MaintainBookmarks\n");
	}
	///*
	//AppendAnnotatedContent_MaintainBookmarks
	//Input Condition: all words in vSentencesVWords are in the cache container
	//*/
	//void AppendAnnotatedContent_MaintainBookmarks(wstring& out, const vector<vector<wstring>>& vSentencesVWords, const vector<wstring>& chunks)
	//{
	//	for (int indSentence = 0; indSentence < vSentencesVWords.size(); ++indSentence)
	//	{
	//		out.append(L"<span class=\"lwtSent\">");

	//		for (int indWord = 0; indWord < vSentencesVWords[indSentence].size(); ++indWord)
	//		{
	//			wstring curWord = WordSansBookmarks(vSentencesVWords[indSentence][indWord]);
	//			unordered_map<wstring,TermRecord>::iterator it = cache->find(curWord);
	//			assert(it != cache->end()); // there should always be a result
	//			if (it->second.nTermsBeyond == 0)
	//			{
	//				AppendWordSpan(out, it->second.wStatus, vSentencesVWords[indSentence][indWord]);
	//				continue;
	//			}

	//			struct mwVals
	//			{
	//				mwVals(wstring wStat, wstring wWC) : wStatus(wStat), wWordCount(wWC) {};
	//				wstring wStatus;
	//				wstring wWordCount;
	//			};
	//			stack<mwVals> stkMWTerms;
	//			// gather a stack of term statuses for tracked multiword terms that start at this word
	//			for (int indTerm = indWord + 1; indTerm < indWord + 9 && indTerm < vSentencesVWords[indSentence].size(); ++indTerm)
	//			{
	//				if (bWithSpaces == true)
	//					curWord.append(L" ");
	//				curWord += WordSansBookmarks(vSentencesVWords[indSentence][indTerm]);

	//				if (usetCacheMWFragments.count(curWord) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
	//					break;

	//				unordered_map<wstring,TermRecord>::iterator it = cache->find(curWord);
	//				if (it != cache->end())
	//					stkMWTerms.push(mwVals(it->second.wStatus, to_wstring(indTerm-indWord+1)));
	//			}

	//			// append span tags for multiword terms, from longest to short
	//			while(stkMWTerms.size() > 0)
	//			{
	//				mwVals mwVCur = stkMWTerms.top();
	//				stkMWTerms.pop();
	//				out.append(L"<span class=\"lwtStat");
	//				out += mwVCur.wStatus;
	//				out.append(L"\"><sup>");
	//				out += mwVCur.wWordCount;
	//				out.append(L"</sup></span>");
	//			}

	//			AppendWordSpan(out, it->second.wStatus, vSentencesVWords[indSentence][indWord]);
	//		}
	//		out.append(L"</span>");
	//	}
	//}
	///*
	//AppendAnnotatedContent_RestoreAtBookmarks
	//Input Condition: all words in vSentencesVWords are in the cache container
	//*/
	//void AppendAnnotatedContent_RestoreAtBookmarks(wstring& out, const vector<vector<wstring>>& vSentencesVWords, const vector<wstring>& chunks)
	//{
	//	for (int indSentence = 0; indSentence < vSentencesVWords.size(); ++indSentence)
	//	{
	//		out.append(L"<span class=\"lwtSent\">");

	//		for (int indWord = 0; indWord < vSentencesVWords[indSentence].size(); ++indWord)
	//		{
	//			wstring curWord = WordSansBookmarks(vSentencesVWords[indSentence][indWord]);
	//			unordered_map<wstring,TermRecord>::iterator it = cache->find(curWord);
	//			assert(it != cache->end()); // there should always be a result
	//			if (it->second.nTermsBeyond == 0)
	//			{
	//				AppendWordSpan(out, it->second.wStatus, Word_RestoreAllBookmarks(vSentencesVWords[indSentence][indWord], chunks));
	//				continue;
	//			}

	//			struct mwVals
	//			{
	//				mwVals(wstring wStat, wstring wWC) : wStatus(wStat), wWordCount(wWC) {};
	//				wstring wStatus;
	//				wstring wWordCount;
	//			};
	//			stack<mwVals> stkMWTerms;
	//			// gather a stack of term statuses for tracked multiword terms that start at this word
	//			for (int indTerm = indWord + 1; indTerm < indWord + 9 && indTerm < vSentencesVWords[indSentence].size(); ++indTerm)
	//			{
	//				if (bWithSpaces == true)
	//					curWord.append(L" ");
	//				curWord += WordSansBookmarks(vSentencesVWords[indSentence][indTerm]);

	//				if (usetCacheMWFragments.count(curWord) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
	//					break;

	//				unordered_map<wstring,TermRecord>::iterator it = cache->find(curWord);
	//				if (it != cache->end())
	//					stkMWTerms.push(mwVals(it->second.wStatus, to_wstring(indTerm-indWord+1)));
	//			}

	//			// append span tags for multiword terms, from longest to short
	//			while(stkMWTerms.size() > 0)
	//			{
	//				mwVals mwVCur = stkMWTerms.top();
	//				stkMWTerms.pop();
	//				out.append(L"<span class=\"lwtStat");
	//				out += mwVCur.wStatus;
	//				out.append(L"\"><sup>");
	//				out += mwVCur.wWordCount;
	//				out.append(L"</sup></span>");
	//			}

	//			AppendWordSpan(out, it->second.wStatus, Word_RestoreAllBookmarks(vSentencesVWords[indSentence][indWord], chunks));
	//		}
	//		out.append(L"</span>");
	//	}
	//}
	void wstring_replaceAll(wstring& out, const wstring& find, const wstring& replace)
	{
		int pos = 0;
		int replaceLen = find.size();
		while (pos != wstring::npos)
		{
			pos = out.find(find, pos);
			if (pos != wstring::npos)
			{
				out.replace(pos, replaceLen, replace);
				pos += replace.size();
			}
		}
	}
	void ReplaceNecessaryCharsWithHTMLEntities(wstring& out)
	{
		//wstring_replaceAll(out, L"&", L"&amp;");
		//wstring_replaceAll(out, L"<", L"&lt;");
	}
	void GetTokensWithoutBookmarks(TokenStruct& tokensWithout, const TokenStruct& tokens)
	{
		for (vector<Token_Sentence>::size_type i = 0; i < tokens.size(); ++i)
		{
			Token_Sentence ts(tokens[i].bDigested);

			for (vector<Token_Sentence>::size_type j = 0; j < tokens[i].size(); ++j)
			{
				ts.push_back(WordSansBookmarks(tokens[i][j]), tokens[i].isWordDigested(j));
			}

			tokensWithout.push_back(ts);
		}
	}
	void GetTokensCanonical(TokenStruct& tokensCanonical, const TokenStruct& tokensWithout)
	{
		for (vector<Token_Sentence>::size_type i = 0; i < tokensWithout.size(); ++i)
		{
			Token_Sentence ts(tokensWithout[i].bDigested);

			for (vector<Token_Sentence>::size_type j = 0; j < tokensWithout[i].size(); ++j)
			{
				ts.push_back(chaj::str::wstring_tolower(tokensWithout[i][j]), tokensWithout[i].isWordDigested(j));
			}

			tokensCanonical.push_back(ts);
		}
	}
	/*	
		!!! External dependencies: Adding New MWSpans relies on the filtering of class global TreeWalker pTW
		SetDocTreeWalker
	*/
	void SetDocTreeWalker(IHTMLDocument2* pDoc, IHTMLElement* pRoot = NULL)
	{
		if (!pDoc) // crash guard for arguments
			return;

		HRESULT hr;
		bool bMustFree = false;

		if (!pRoot)
		{
			pDoc->get_body(&pRoot);
			if (!pRoot) return;
			bMustFree = true;
		}

		IDocumentTraversal* pDT = NULL;
		hr = pDoc->QueryInterface(IID_IDocumentTraversal, (void**)&pDT);
		if (!pDT)
		{
			if (bMustFree)
			{

				pRoot->Release();
				return;
			}
		}

		VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = pFullFilter;
		pDT->createTreeWalker(pRoot, SHOW_ELEMENT, &varFilter, VARIANT_TRUE, &pTW);
	}
	void ReparsePage(IHTMLDocument2* pDoc)
	{
		wstring theNewsBodyContent(L"<style>span.lwtStat0 {background-color: #ADDFFF} span.lwtStat1 {background-color: #F5B8A9} span.lwtStat2 {background-color: #F5CCA9} span.lwtStat3 {background-color: #F5E1A9} span.lwtStat4 {background-color: #F5F3A9} span.lwtStat5 {background-color: #DDFFDD} span.lwtStat99 {background-color: #F8F8F8;border-bottom: solid 2px #CCFFCC} span.lwtStat98 {background-color: #F8F8F8;border-bottom: dashed 1px #000000}</style>");
		
		AppendAnnotatedContent_MaintainBookmarks(theNewsBodyContent, tokens, tknSansBookmarks, tknCanonical);

		ReplaceNecessaryCharsWithHTMLEntities(theNewsBodyContent);

		RestoreAllBookmarks(theNewsBodyContent, chunks);

		// replace body text with added annotations
		_bstr_t bNewBodyContent = theNewsBodyContent.c_str();
		IHTMLElement* pBody = GetBodyFromDoc(pDoc);
		HRESULT hr = pBody->put_innerHTML(bNewBodyContent.GetBSTR());
		pBody->Release();
	}
	void ExpandChunkEntities(vector<wstring>& chunks)
	{
		for (vector<wstring>::size_type i = 0; i < chunks.size(); ++i)
			ReplaceHTMLEntitiesWithChars(chunks[i]);
	}
	void ParsePage(IHTMLDocument2* pDoc)
	{
		if (!pDoc) // crash guard for arguments
			return;

		IHTMLElement* pBody = NULL;
		wstring wsBodyContent = L"";
		GetBodyContent(wsBodyContent, pBody, pDoc);	
		if (pBody == NULL)
			return;

		// separate html into chunks, each chunk fully one of intra-tag content or extra-tag content:  |<p>|This is |<u>|such|</u>| a good book!|</p>|
		chunks.clear();
		PopulateChunks(chunks, wsBodyContent);

		// condense non visible data down to placeholders
		wstring wVisibleBody = GetVisibleBodyWithTagBookmarks(chunks);

		// handle html entities such as &lt; and &#62, so that word parsing parses what you see and not how it's represented
		// this is done after chunking so chunking has less issues with embedded '<' chars, and it's done
		// after non-visible areas (like scripts) have been bookmarked to avoid altering programmatic features we don't see
		// anyway
		ReplaceHTMLEntitiesWithChars(wVisibleBody);

		tokens.vSentences.clear();
		tknSansBookmarks.vSentences.clear();
		tknCanonical.vSentences.clear();
		ParseIntoSentencesAsTokenStruct(tokens, wVisibleBody);
		assert(TokensReassembledProperly(tokens, wVisibleBody));
		GetTokensWithoutBookmarks(tknSansBookmarks, tokens);
		GetTokensCanonical(tknCanonical, tknSansBookmarks);

		UpdateCaches(tknCanonical);
		ReconstructDoc(tokens, tknSansBookmarks, tknCanonical, pDoc, pBody);	

#ifdef _DEBUG
		IHTMLElement* pBody2 = NULL;
		wstring wsBodyContent2 = L"";
		GetBodyContent(wsBodyContent2, pBody2, pDoc);
		if (pBody2)
			pBody2->Release();
#endif

		pBody->Release();
	}
	void AppendTermDivRec(IHTMLElement* pBody, const wstring& wTermCanonical, const TermRecord& rec)
	{
		wstring out;
		out.append(L"<div id=\"lwt");
		out.append(to_wstring(rec.uIdent));
		out.append(L"\" lwtstat=\"");
		out.append(rec.wStatus);
		out.append(L"\" lwtterm=\"");
		out.append(wTermCanonical);
		out.append(L"\" lwttrans=\"");
		out.append(rec.wTranslation);
		out.append(L"\" lwtrom=\"");
		out.append(rec.wRomanization);
		out.append(L"\" />");
		chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
	}
	void AppendTermDivRecs(IHTMLElement* pBody)
	{
		unordered_set<wstring> termSet;
		GetSetOfTerms(termSet, tknCanonical);

		for (auto it = termSet.cbegin(); it != termSet.cend(); ++it)
		{
			AppendTermDivRec(pBody, *it, cache->find(*it)->second);
		}
	}
	wstring GetDropdownHTML_TableSet()
	{
		wstring wResult;
		wResult.append(L"<option value=\"\"");
		if (wTableSetPrefix == L"")
			wResult.append(L" selected=\"selected\"");
		wResult.append(L">Default Table Set</option>");

		vector<wstring> vPrefixes;
		GetDropdownPrefixList(vPrefixes);

		for (wstring& set : vPrefixes)
		{
			wResult.append(L"<option value=\"");
			wResult.append(set);
			wResult.append(L"\"");
			if (wTableSetPrefix.find(set) == 0 && wTableSetPrefix.size() == set.size() + 1)
				wResult.append(L" selected=\"selected\"");
			wResult.append(L">");
			wResult.append(set);
			wResult.append(L"</option>");
		}

		return wResult;
	}
	wstring GetDropdownHTML_LanguageChoice()
	{
		wstring wResult = L"";
		wstring wQuery = L"select LgID, LgName from ";
		wQuery.append(wTableSetPrefix);
		wQuery.append(L"languages");
		if (this->DoExecuteDirect(L"3394uhfljh", wQuery))
		{
			while (tStmt.fetch_next())
			{
				wstring curID = tStmt.field(1).as_string();
				wResult.append(L"<option value=\"");
				wResult.append(curID);
				wResult.append(L"\"");
				if (wLgID == curID)
					wResult.append(L" selected=\"selected\"");
				wResult.append(L">");
				wResult.append(tStmt.field(2).as_string());
				wResult.append(L"</option>");
			}
			tStmt.free_results();
		}
		
		return wResult;
	}
	void AppendHtmlBlocks(IHTMLDocument2* pDoc)
	{
		if (!pDoc)
			return;

		IHTMLElement* pBody = GetBodyFromDoc(pDoc); SmartCOMRelease scBody(pBody);
		if (!pBody)
			return;

		wstring out;
		out.append(
		L"<div id=\"lwtinlinestat\" style=\"display:none;position:absolute;\" onmouseout=\"lwtdivmout(event);\">"
		L"<span id=\"lwtTermTrans\" style=\"display:inline-block\">&nbsp;</span><br />"
		L"<span class=\"lwtStat1\" id=\"lwtsetstat1\" lwtstat=\"1\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center\" lwtstatchange=\"true\">1</span>"
		L"<span class=\"lwtStat2\" id=\"lwtsetstat2\" lwtstat=\"2\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">2</span>"
		L"<span class=\"lwtStat3\" id=\"lwtsetstat3\" lwtstat=\"3\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">3</span>"
		L"<span class=\"lwtStat4\" id=\"lwtsetstat4\" lwtstat=\"4\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">4</span>"
		L"<span class=\"lwtStat5\" id=\"lwtsetstat5\" lwtstat=\"5\" style=\"border:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">5</span><br />"
		L"<span class=\"lwtStat99\" id=\"lwtsetstat99\" lwtstat=\"99\" style=\"border-left:solid black 1px;border-bottom: solid 1px #CCFFCC;display:inline-block;width:26px;text-align:center;\" lwtstatchange=\"true\" title=\"Well Known\">W</span>"
		L"<span class=\"lwtStat98\" id=\"lwtsetstat98\" lwtstat=\"98\" style=\"border-left:solid black 1px;display:inline-block;width:26px;text-align:center;\" lwtstatchange=\"true\" title=\"Ignore\">Ig</span>"
		L"<span class=\"lwtStat0\" id=\"lwtsetstat0\" lwtstat=\"0\" style=\"border-left:solid black 1px;border-right:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:25px;text-align:center;\" lwtstatchange=\"true\" title=\"Untrack\">Un</span><br />"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 1\"><a id=\"lwtextrapdict1\" href=\"\">D1</a></span>"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 2\"><a id=\"lwtextrapdict2\" href=\"\">D2</a></span>"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:25px;text-align:center;background-color:white;\" title=\"Google Translate Term\"><a id=\"lwtextrapgoogletrans\" href=\"\">GT</a></span><br />"
		L"<span id=\"lwtmwstart\" style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:79px;text-align:center;background-color:white;\" onclick=\"beginMW();\" mwbuffer=\"true\" title=\"Begin Multiple Word Term\">Multi</span><br />"
		L"<span onclick=\"lwtStartEdit();\" style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:79px;text-align:center;background-color:white;\">Edit</span>"
		L"</div>"
		L"<div id=\"lwtlasthovered\" style=\"display:none;\" lwtterm=\"\"></div>"
		);

		out.append(
		L"<div id=\"lwtInlineMWEndPopup\" style=\"display:none;position:absolute;\" onmouseout=\"lwtdivmout(event);\">"
		L"<span id=\"lwtTermTrans2\">&nbsp;</span><br />"
		L"<span class=\"lwtStat1\" lwtstat=\"1\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center\" onclick=\"captureMW(49);\">1</span>"
		L"<span class=\"lwtStat2\" lwtstat=\"2\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" onclick=\"captureMW(50);\">2</span>"
		L"<span class=\"lwtStat3\" lwtstat=\"3\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" onclick=\"captureMW(51);\">3</span>"
		L"<span class=\"lwtStat4\" lwtstat=\"4\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" onclick=\"captureMW(52);\">4</span>"
		L"<span class=\"lwtStat5\" lwtstat=\"5\" style=\"border:solid black 1px;display:inline-block;width:15px;text-align:center;\" onclick=\"captureMW(53);\">5</span><br />"
		L"<span class=\"lwtStat99\" lwtstat=\"99\" style=\"border-left:solid black 1px;border-bottom: solid 1px #CCFFCC;display:inline-block;width:39px;text-align:center;\" onclick=\"captureMW(87);\" title=\"Well Known\">Well</span>"
		L"<span class=\"lwtStat98\" lwtstat=\"98\" style=\"border-left:solid black 1px;border-right:solid black 1px;display:inline-block;width:39px;text-align:center;\" onclick=\"captureMW(73);\" title=\"Ignore\">Igno</span><br />"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 1\"><a id=\"lwtextrapdict1_2\" href=\"javascript:multiWordDict('D1');\">D1</a></span>"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 2\"><a id=\"lwtextrapdict2_2\" href=\"javascript:multiWordDict('D2');\">D2</a></span>"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:25px;text-align:center;background-color:white;\" title=\"Google Translate Term\"><a id=\"lwtextrapgoogletrans_2\" href=\"javascript:multiWordDict('DG');\">GT</a></span><br />"
		L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:79px;text-align:center;background-color:white;\" onclick=\"cancelMW();\" mwbuffer=\"true\" title=\"Abort Multiple Word Term\">Abort</span>"
		L"</div>"
		L"<div style=\"display:none;\"><a id=\"lwtMultiDictLink\" href=\"\"></a></div>"
		);

		out.append(
		L"<div id=\"lwtterminfo\" style=\"position:absolute;left:0;top:0;width:100%;background-color:white;z-index:5;display:none;\">"
		L"<div style=\"width:50%;height:100%;float:left\">"
		L"&nbsp;Definition of <span id=\"lwtSetTermLabel\"></span>:<br />"
		L"<textarea id=\"lwtshowtrans\" style=\"width:90%;display:inline;\" rows=\"3\"></textarea><br />"
		L"&nbsp;Romanization:<br />"
		L"<textarea id=\"lwtshowrom\" style=\"font-family:'arial';width:90%;display:inline;\" rows=\"1\"></textarea><br />"
		L"<div style=\"margin:0px auto;width:200px;\"><button type=\"button\" style=\"\" lwtAction=\"lwtUpdateTermInfo\">Update</button>&nbsp;&nbsp"
		L"<button type=\"button\" style=\"\" onclick=\"closeEditBox();\">Close</button></div>"
		L"</div><div style=\"width:50%;float:right;\">"
		L"<iframe id=\"lwtiframe\" name=\"lwtiframe\" src=\"\" style=\"width:100%\"></iframe>"
		L"</div>"
		L"</div>"
		L"<div id=\"lwtcurinfoterm\" style=\"display:none;\" lwtterm=\"\"></div>"
		);

		out.append(
		L"<div id=\"lwtPopupInfo\" style=\"font-family:'arial';position:absolute;left:0;top:0;width:100%;z-index:100;display:none;\">"
		L"<div id=\"lwtPopupTrans\" style=\"border: solid black 1px;background-color:white;clear:both;\"></div>"
		L"<div id=\"lwtPopupRom\" style=\"border-bottom:solid black 1px;border-left:solid black 1px;border-right:solid black 1px;background-color:white;clear:both\"></div>"
		L"</div>"
		);

		wstring wDropdown = GetDropdownHTML_LanguageChoice();
		wstring wDropdown2 = GetDropdownHTML_TableSet();

		out.append(
			L"<div id=\"lwtSettings\" style=\"position:absolute;left:0;top:0;width:100%;background-color:white;padding:5px;display:none;z-index:1000;\">"
			L"Choose language: <select id=\"lwtLangDropdown\" onchange=\"lwtChangeLang(this);\">");
		out.append(wDropdown.c_str());
		out.append(
			L"</select><br /><br />"
			L"Choose Table Set: <select id=\"lwtTableSetDropdown\" onchange=\"lwtChangeTableSet(this);\">");
		out.append(wDropdown2.c_str());
		out.append(
			L"</select><br />"
			L"<br />"
			L"<button type=\"button\" onclick=\"document.getElementById('lwtSettings').style.display = 'none';\">Close Settings</button>"
			L"</div>"
			L"<div id=\"lwtCurLangChoice\" style=\"display:none;\" value=\"\" lwtAction=\"changeLang\"></div>"
			L"<div id=\"lwtBhoCommand\" style=\"display:none;\" value=\"\" onclick=\"lwtExecBhoCommand()\"></div>"
			L"<div id=\"lwtJSCommand\" style=\"display:none;\" value=\"\" lwtAction=\"\"></div>"
			);

		out.append(L"<div id=\"lwtdict1\" style=\"display:none;\" src=\"");
		out.append(wstrDict1);
		out.append(L"\"></div>");

		out.append(L"<div id=\"lwtdict2\" style=\"display:none;\" src=\"");
		out.append(wstrDict2);
		out.append(L"\"></div>");

		out.append(L"<div id=\"lwtgoogletrans\" style=\"display:none;\" src=\"");
		out.append(wstrGoogleTrans);
		out.append(L"\"></div>");

		out.append(L"<div id=\"lwtwithspaces\" style=\"display:none;\" value=\"");
		if (bWithSpaces)
			out.append(L"yes");
		else
			out.append(L"no");
		out.append(L"\"></div>");

		out.append(L"<div id=\"lwtSetupLink\" style=\"display:none;\" onclick=\"lwtSetup();\"></div>");

		HRESULT hr = chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
		assert(SUCCEEDED(hr));

		out = L"<div id=\"lwtIntroDiv\" style=\"background-color:white\"></div>";
		hr = chaj::DOM::AppendHTMLAfterBegin(out, pBody);
		assert(SUCCEEDED(hr));
	}
	void GetSetOfTerms(unordered_set<wstring>& out, const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling GetSetOfTerms\n");
		for (vector<Token_Sentence>::size_type i = 0; i < tknCanonical.size(); ++i)
		{
			if (tknCanonical[i].bDigested)
			{
				for (vector<Token_Sentence>::size_type j = 0; j < tknCanonical[i].size(); ++j)
				{
					if (tknCanonical[i].isWordDigested(j))
					{
						unordered_map<wstring,TermRecord>::const_iterator it = cache->find(tknCanonical[i][j]);
						assert(it != cache->end());
						out.insert(tknCanonical[i][j]);

						if (it->second.nTermsBeyond != 0) //the term might participate in multiword terms
						{
							wstring curTerm = tknCanonical[i][j];
							if (usetCacheMWFragments.count(curTerm) != 0) //this terms starts some MWTerm
							{
								int nSkipCount = 0;
								// gather a stack of term statuses for tracked multiword terms that start at this word
								for (vector<Token_Word>::size_type k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
								{
									if (tokens[i].isWordDigested(k) == false)
									{
										++nSkipCount;
										continue;
									}

									if (bWithSpaces == true)
										curTerm.append(L" ");
									curTerm += tknCanonical[i][k];

									if (usetCacheMWFragments.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
										break;

									unordered_map<wstring,TermRecord>::const_iterator it = cache->find(curTerm);
									if (it != cache->end())
										out.insert(curTerm);
								}
							}
						}
					}
				}
			}
		}
		TRACE(L"%s", L"Leaving GetSetOfTerms\n");
	}
	void AppendCss(IHTMLDocument2* pDoc)
	{
		HRESULT hr = AppendStylesToDoc(this->wCss, pDoc);
		assert(SUCCEEDED(hr));
	}
	void dhr(HRESULT hr) //(D)ebug (HR)esult
	{
		assert(SUCCEEDED(hr));
	}
	void ReconstructDoc(const TokenStruct& tokens, const TokenStruct& tknWithout, const TokenStruct& tknCanonical, IHTMLDocument2* pDoc, IHTMLElement* pBody)
	{
		HRESULT hr;
		VARIANT_BOOL vBool;
		wstring out;
		int lwtid = 1;

		AppendTermDivRecs(pBody);
		AppendHtmlBlocks(pDoc);
		/* specifically last after all elements placed */ AppendJavascript(pDoc);

		IHTMLElement* pSetupLink = chaj::DOM::GetElementFromId(L"lwtSetupLink", pDoc);
		assert(pSetupLink);
		dhr(pSetupLink->click());
		pSetupLink->Release();

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		if (!pRange)
			return;

		wstring wstrText = HTMLTxtRange_get_text(pRange);
		wstring wstrHTMLText = HTMLTxtRange_htmlText(pRange);

		AppendCss(pDoc);

		_bstr_t bSentenceMarker(L"<span class=\"lwtSent\" lwtsent=\"lwtsent\"></span>");
		BSTR bstrSentenceMarker = SysAllocString(L"<span class=\"lwtSent\" lwtsent=\"lwtsent\"></span>");
		dhr(chaj::DOM::TxtRange_CollapseToBegin(pRange));

		for (decltype(tokens.size()) i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			dhr(pRange->pasteHTML(bstrSentenceMarker));
			
			for (decltype(tokens[i].size()) j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				BSTR bWord = SysAllocString(tknWithout[i][j].c_str());
				pRange->findText(bWord, 1000, 0, &vBool);
				SysFreeString(bWord);
				if (vBool == VARIANT_FALSE)
				{
					TRACE(L"Warning: Could not locate word to replace it: %s\n", tknWithout[i][j].c_str());
					continue;
				}

				cache_cit it = cache->find(tknCanonical[i][j]); assert(it != cache->end()); // there should always be a result

				if (it->second.nTermsBeyond != 0) //the term might participate in multiword terms
				{
					wstring curTerm = tknCanonical[i][j];
					if (usetCacheMWFragments.count(curTerm) != 0) //this terms starts some MWTerm
					{
						stack<mwVals> stkMWTerms;
						int nSkipCount = 0;
						// gather a stack of term statuses for tracked multiword terms that start at this word
						for (decltype(tokens[i].size()) k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
						{
							if (tokens[i].isWordDigested(k) == false)
							{
								++nSkipCount;
								continue;
							}

							if (bWithSpaces == true)
								curTerm.append(L" ");
							curTerm += tknCanonical[i][k];

							if (usetCacheMWFragments.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
								break;

							cache_cit it = cache->cfind(curTerm);
							if (it != cache->cend())
								stkMWTerms.push(mwVals(&(it->second), to_wstring(k-j+1-nSkipCount), curTerm));
						}

						// append span tags for multiword terms, from longest to short
						while(stkMWTerms.size() > 0)
						{
							mwVals mwVCur = stkMWTerms.top();
							stkMWTerms.pop();
							AppendMWSpan(out, mwVCur.wMWTerm, mwVCur.pRec, mwVCur.wWordCount);
						}
					}
				}

				if (tokens[i][j].find(tknWithout[i][j]) != wstring::npos)
				{
					AppendSoloWordSpan(out, tknWithout[i][j], tknCanonical[i][j], &(it->second));
					BSTR bNew = SysAllocString(out.c_str());
					hr = pRange->pasteHTML(bNew);
					assert(SUCCEEDED(hr));
					SysFreeString(bNew);
				}
				else
				{
					dhr(chaj::DOM::TxtRange_CollapseToEnd(pRange));
				}
				out.clear();
			}
		}

		dhr(pRange->pasteHTML(bstrSentenceMarker));

		SysFreeString(bstrSentenceMarker);
	}
	bool TokensReassembledProperly(const TokenStruct& ts, const wstring& orig)
	{
		wstring result;

		for(vector<Token_Sentence>::size_type i = 0; i < ts.size(); ++i)
		{
			for(vector<Token_Word>::size_type j = 0; j < ts[i].size(); ++j)
			{
				result.append(ts[i][j]);
			}
		}
		return (wcscmp(orig.c_str(), result.c_str()) == 0);
	}
	HRESULT __stdcall SetSite(IUnknown* pUnkSite)
	{
		HRESULT hr;

		if (pUnkSite == NULL || pSite != NULL)
		{
			if (pCP != NULL)
			{
				if (dwCookie != NULL)
				{
					hr = pCP->Unadvise(dwCookie);
				}
				pCP->Release();
			}
			if (pSite != NULL)
				pSite->Release();
			if (pBrowser != NULL)
				pBrowser->Release();
			if (pDispBrowser != NULL)
				pDispBrowser->Release();
			if (pTW != NULL)
				pTW->Release();
			return S_OK;
		}

		pUnkSite->AddRef();
		if (pCP != NULL)
		{
			if (dwCookie != NULL)
			{
				hr = pCP->Unadvise(dwCookie);
			}	
			pCP->Release();
		}
		if (pSite != NULL)
			pSite->Release();
		if (pBrowser != NULL)
			pBrowser->Release();
		if (pDispBrowser != NULL)
			pDispBrowser->Release();
		pSite = pUnkSite;

		hr = pUnkSite->QueryInterface(IID_IWebBrowser2, (void**)&pBrowser);
		if (FAILED(hr))
		{
			mb(_T("IWebBrowser2 fail"), _T("1495hfhoksj"));
			pSite->Release();
			pSite = NULL;
			return E_UNEXPECTED;
		}

		pDispBrowser = GetAlternateInterface<IWebBrowser2, IDispatch>(pBrowser);

		IConnectionPointContainer* pCPC;
		hr = pUnkSite->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
		if (FAILED(hr))
		{
			mb(_T("ICPC fail"), _T("1505hdbvllkjd"));
			pBrowser->Release();
			pSite->Release();
			pSite = NULL;
			return E_UNEXPECTED;
		}

		pBrowser->get_HWND((SHANDLE_PTR*)&hBrowser);
		hMBParent = hBrowser;

		hr = pCPC->FindConnectionPoint(DIID_DWebBrowserEvents2, &pCP);
		pCPC->Release();
		if (FAILED(hr))
		{
			mb(_T("FindCP fail"), _T("465gasydg"));
			pBrowser->Release();
			pSite->Release();
			pSite = NULL;
			hBrowser = NULL;
			hBrowser = NULL;
			return hr;
		}

		hr = pCP->Advise(reinterpret_cast<IDispatch*>(this), &dwCookie);
		if (FAILED(hr))
		{
			mb(_T("CPAdvise fail"), _T("472sfsxc"));
			pBrowser->Release();
			pSite->Release();
			pCP->Release();
			pSite = NULL;
			hBrowser = NULL;
			hBrowser = NULL;
			return hr;
		}

		return hr;
	}
	HRESULT __stdcall GetSite(REFIID riid, void** ppvSite)
	{
		if (pSite != NULL)
		{
			IUnknown* iThis;
			HRESULT hr = pSite->QueryInterface(riid, (void**)&iThis);
			if (FAILED(hr))
				return E_NOINTERFACE;
			else
			{
				*ppvSite = pSite;
				return S_OK;
			}
		}
		else
			return E_FAIL;
	}
	IDOMTreeWalker* GetDocumentTree(IHTMLDocument2* pDoc)
	{
		if (!pDoc) // crash guard for arguments
			return nullptr;

		HRESULT hr;

		IHTMLElement* pRoot = nullptr;
		pDoc->get_body(&pRoot); SmartCOMRelease scRoot(pRoot);
		if (!pRoot)
			return nullptr;

		IDocumentTraversal* pDT = nullptr;
		hr = pDoc->QueryInterface(IID_IDocumentTraversal, (void**)&pDT); SmartCOMRelease scDT(pDT);
		if (!pDT)
			return nullptr;

		IDOMTreeWalker* pTW = nullptr;
		VARIANT varNull; varNull.vt = VT_NULL;
		hr = pDT->createTreeWalker(pRoot, SHOW_ELEMENT | SHOW_TEXT, &varNull, VARIANT_TRUE, &pTW); // allow requestor to Release
		return pTW;
	}
	cache_it CacheTermAndReturnIt(const wstring& wWord) // optimizable
	{
		list<wstring> cWord;
		cWord.push_back(wWord);
		EnsureRecordEntryForEachWord(cWord);
		return cache->find(wWord);
	}
	/*
		Function ParseTextNode_GetReplacementNode

		This function is heavily interrelated with ParseTextNode, in that together they manage reference counts in a non-independent way

		Return value:
			--(on error): returns nullptr
			--(otherwise): pointer to node to replace
		Exit conditions:
			--the pText that came in was released
			--If return is nullptr, no new pText has been obtained
			--If return is non-nullptr, a new Ptext had been obtain if bFurtherText, otherwise one has ~not~ been obtained
	*/
	wstring GetText_Helper__ParseTextNode_AllAtOnce3(IHTMLDOMTextNode* pText)
	{
		wstring wText;
		_bstr_t bText;
		pText->get_data(bText.GetAddress());
		if (!bText.length())
			return wText;

		wText = bText;
		return wText;
	}
	void ParseTextNode_AllAtOnce(IHTMLDOMTextNode* pText, IHTMLDocument2* pDoc, unordered_set<wstring>& usetPageTerms)
	{
		if (!pText || !pDoc)
			return;

		wstring wText = GetText_Helper__ParseTextNode_AllAtOnce3(pText);
		if (wText.empty())
			return;

		wstring wOuterHTML;
		unsigned int lastPosition = 0;
		unsigned int textLength = wText.length();

		IHTMLElement *pNew = CreateElement(pDoc, L"span"); SmartCOMRelease scNew(pNew);
		if (!pNew)
			return;

		IHTMLDOMNode* pNewNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pNew); SmartCOMRelease scNewNode(pNewNode);
		IHTMLDOMNode* pTextNode = GetAlternateInterface<IHTMLDOMTextNode,IHTMLDOMNode>(pText); SmartCOMRelease scTextNode(pTextNode);
		IHTMLElement* pBody = GetBodyFromDoc(pDoc); SmartCOMRelease scBody(pBody);
		if (!pBody || !pNewNode || !pTextNode)
			return;

		wstring regPtn = L"[" + WordChars + L"]";
		if (bWithSpaces)
			regPtn.append(L"+");

		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator regit(wText.begin(), wText.end(), wrgx);
		wregex_iterator rend;

		vector<wstring> words;
		wstring termSpacer = L"~~!@~!XXSpacerLwtBho";

		// get list of words with spacers at breaks between possible multi-word terms
		while (regit != rend)
		{
			if (regit->position() != lastPosition)
			{
				wstring skipped = wstring(wText.begin()+lastPosition, wText.begin()+regit->position());
				if (!bWithSpaces ||
					(bWithSpaces && !std::all_of(skipped.begin(), skipped.end(), [](wchar_t in){return in == L' ';})))
					words.push_back(termSpacer);
			}

			lastPosition = regit->position() + regit->length();
			words.push_back(regit->str());
			regit++;
		}
		if (words.empty())
			return; // no words found, just return

		// create a stack of records for any multi-word term records and add their div recs to the page
		lastPosition = 0; // reset position
		int curPosition = 0;
		for(unsigned int i = 0; i < words.size(); ++i)
		{
			if (words[i] == termSpacer)
				continue;

			stack<tuple<wstring,TermRecord,int>> mwTerms;
			wstring firstWord = chaj::str::wstring_tolower(words[i]);
			wstring term = firstWord;

			if (usetCacheMWFragments.find(term) != usetCacheMWFragments.end())
				// some mwTerms begins with this word
			{
				for(unsigned int j = 1; i+j < LWT_MAX_MWTERM_LENGTH && i+j < words.size(); ++j)
				{
					if (words[i+j] == termSpacer)
						break;

					if (bWithSpaces)
						term += L" ";
					term += chaj::str::wstring_tolower(words[i+j]);

					if (usetCacheMWFragments.find(term) == usetCacheMWFragments.end())
						// to mwTerms begin with this substring
						break;

					cache_cit it = cache->cfind(term);
					if (TermCached(it))
					{
						mwTerms.push(make_tuple(term, it->second, j+1));
						if (usetPageTerms.find(term) == usetPageTerms.end())
						{
							AppendTermDivRec(pBody, term, it->second);
							usetPageTerms.insert(term);
						}
					}
				}
			}

			curPosition = wText.find(words[i], lastPosition);
			assert(curPosition != wstring::npos); // we already found the word, it should still be there

			if (lastPosition != curPosition)
				wOuterHTML.append(wstring(wText.begin()+lastPosition, wText.begin()+curPosition));

			lastPosition = curPosition + firstWord.size();

			while (mwTerms.size() > 0)
			{
				tuple<wstring, TermRecord, int> current = mwTerms.top(); mwTerms.pop();

				AppendMWSpan(wOuterHTML, get<0>(current), &get<1>(current), to_wstring(get<2>(current)));
			}

			cache_cit it = cache->cfind(firstWord);
			if (TermNotCached(it))
			{
				it = CacheTermAndReturnIt(firstWord);
			}
			assert(it != cache->end());
			AppendSoloWordSpan(wOuterHTML, words[i], firstWord, &it->second);
			AppendTermDivRec(pBody, firstWord, it->second);
		}

		if (lastPosition != wText.size()) // non-word text beyond last found word
			wOuterHTML.append(wText.begin()+lastPosition, wText.end()); // append all remaining

		HRESULT hr = pTextNode->replaceNode(pNewNode, &pTextNode);
		if (SUCCEEDED(hr))
			hr = chaj::DOM::SetElementOuterHTML(pNew, wOuterHTML);
	}
	void RoughInsert(list<wstring>& cWords, int nAtATime)
	{
		CacheDBHitsWithListRemove(cWords, nAtATime);
		if (bShuttingDown)
			return;

		// cache misses with status 0
		for(wstring wTerm : cWords)
		{
			if (bShuttingDown)
				return;
			cache->insert(unordered_map<wstring,TermRecord>::value_type(wTerm, TermRecord(L"0")));
		}
	}
	void RoughInsertList(forward_list<wstring>& cWords, IHTMLElement* pBody, unordered_set<wstring>& pageDivRecs)
	{	
		wstring	wInList;

		// create string of uncached terms
		for (auto iter = cWords.begin(); iter != cWords.end(); ++iter)
		{
			cache_it it = cache->find(*iter);
			if (TermNotCached(it))
			{
				if (!wInList.empty())
					wInList.append(L",");

				wInList.append(L"'");
				wInList.append(EscapeSQLQueryValue(*iter));
				wInList.append(L"'");
			}
		}

		// cache uncached tracked terms
		if (!wInList.empty())
		{
			wstring wQuery;
			wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from ");
			wQuery.append(wTableSetPrefix);
			wQuery.append(L"words where WoLgID = ");
			wQuery.append(wLgID);
			wQuery.append(L" AND WoTextLC in (");
			wQuery.append(wInList);
			wQuery.append(L")");

			EnterCriticalSection(&CS_UseDBConn);
			DoExecuteDirect(_T("1612isjdlfij"), wQuery);
			while (tStmt.fetch_next())
			{
				wstring wLC = tStmt.field(1).as_string();
				TermRecord rec(tStmt.field(2).as_string());
				rec.wRomanization = tStmt.field(3).as_string();
				rec.wTranslation = tStmt.field(4).as_string();
				cache->insert(unordered_map<wstring,TermRecord>::value_type(wLC, rec));
			}
			tStmt.free_results();
			LeaveCriticalSection(&CS_UseDBConn);
		}

		// at this point all terms in list that are tracked are cached

		if (bShuttingDown)
			return;

		// add all necessary DivRecs to the page
		for(wstring wTerm : cWords)
		{
			cache_cit it = cache->find(wTerm);
			if (TermNotCached(it))
			{
				cache->insert(unordered_map<wstring,TermRecord>::value_type(wTerm, TermRecord(L"0")));
				AppendTermDivRec(pBody, wTerm, TermRecord(L"0"));
				pageDivRecs.insert(wTerm);
			}
			else if (pageDivRecs.find(wTerm) == pageDivRecs.end())
			{
				AppendTermDivRec(pBody, wTerm, it->second);
				pageDivRecs.insert(wTerm);
			}
			
			if (bShuttingDown)
				return;
		}
	}
	void Thread_MWHighlight(LPSTREAM pDocStream, forward_list<wstring>& cWords)
	{
		if (!pDocStream)
			return;

		IHTMLDocument2* pDoc = nullptr;
		HRESULT hr = CoGetInterfaceAndReleaseStream(pDocStream, IID_IHTMLDocument2, reinterpret_cast<LPVOID*>(&pDocStream));
		SmartCOMRelease scDoc(pDoc, true); // AddRef and schedule Release
		if (FAILED(hr) || !pDoc)
			return;

		forward_list<wstring> possTerms;
		
		for (auto iter = cWords.begin(); iter != cWords.end(); ++iter)
		{
			auto subIter(iter);
			++subIter;
			unsigned int additionalTerms = 1;
			wstring wCumTerm = *iter;
			while (subIter != cWords.end() && additionalTerms < LWT_MAX_MWTERM_LENGTH)
			{
				wCumTerm += *subIter;
				possTerms.push_front(wCumTerm);
				++additionalTerms;
				++subIter;
			}
		}
	}
		void Thread_CachePageMWTerms(LPSTREAM pDocStream, shared_ptr<vector<wstring>> sp_vWords)
	{
		if (!pDocStream || !sp_vWords)
			return;

		const int MAX_MYSQL_INLIST_LEN = 500;

		IHTMLDocument2* pDoc = nullptr;
		HRESULT hr = CoGetInterfaceAndReleaseStream(pDocStream, IID_IHTMLDocument2, reinterpret_cast<LPVOID*>(&pDoc));
		SmartCOMRelease scDoc(pDoc, true); // AddRef and schedule Release
		if (FAILED(hr) || !pDoc)
			return;

		for (int i = 0; i < sp_vWords->size(); i += MAX_MYSQL_INLIST_LEN)
		{
			wstring	wInList;

			// create buffer of uncached terms
			for (int j = 0; j < MAX_MYSQL_INLIST_LEN && i+j < sp_vWords->size(); ++j)
			{
				wstring wRoot = (*sp_vWords)[i+j];
				wstring wTerm = wRoot;
				for (unsigned int k = 1; k <= 8 && i+j+k < (*sp_vWords).size(); ++k)
				{
					if (this->bWithSpaces)
						wTerm += L" ";

					wTerm += (*sp_vWords)[i+j+k];

					if (!wInList.empty())
						wInList.append(L",");

					wInList.append(L"'");
					wInList.append(EscapeSQLQueryValue(wTerm));
					wInList.append(L"'");
				}
			}

			// cache uncached tracked terms
			if (!wInList.empty())
			{
				wstring wQuery;
				wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from ");
				wQuery.append(wTableSetPrefix);
				wQuery.append(L"words where WoLgID = ");
				wQuery.append(wLgID);
				wQuery.append(L" AND WoTextLC in (");
				wQuery.append(wInList);
				wQuery.append(L")");

				EnterCriticalSection(&CS_UseDBConn);
				DoExecuteDirect(_T("1612isjdlfij"), wQuery);
				while (tStmt.fetch_next())
				{
					wstring wLC = tStmt.field(1).as_string();
					TermRecord rec(tStmt.field(2).as_string());
					rec.wRomanization = tStmt.field(3).as_string();
					rec.wTranslation = tStmt.field(4).as_string();
					cache->insert(unordered_map<wstring,TermRecord>::value_type(wLC, rec));
					UpdateCacheMWFragments(wLC);
				}
				tStmt.free_results();
				LeaveCriticalSection(&CS_UseDBConn);
			}
			// at this point all terms in list that are tracked are cached

			if (bShuttingDown)
				return;
		}
	}
	void Helper_Notify_Thread_CachePageWords(bool* pbHeadStartComplete, condition_variable* p_cv)
	{
		{
			mutex m;
			std::lock_guard<mutex> lk(m);
			*pbHeadStartComplete = true;
		}
		p_cv->notify_all();
	}
	void Thread_CachePageWords(LPSTREAM pDocStream, shared_ptr<vector<wstring>> sp_vWords, condition_variable*  p_cv, bool* pbHeadStartComplete)
	{
		if (!pDocStream || !sp_vWords)
			return;

		const int MAX_MYSQL_INLIST_LEN = 500;

		IHTMLDocument2* pDoc = nullptr;
		HRESULT hr = CoGetInterfaceAndReleaseStream(pDocStream, IID_IHTMLDocument2, reinterpret_cast<LPVOID*>(&pDoc));
		SmartCOMRelease scDoc(pDoc, true); // AddRef and schedule Release
		if (FAILED(hr) || !pDoc)
			return;

		if ((*sp_vWords).size() == 0)
		{
			Helper_Notify_Thread_CachePageWords(pbHeadStartComplete, p_cv);
		}
		else
		{
			for (int i = 0; i < (*sp_vWords).size(); i += MAX_MYSQL_INLIST_LEN)
			{
				list<wstring> cWords;
				wstring	wInList;

				// create buffer of uncached terms
				for (int j = 0; j < MAX_MYSQL_INLIST_LEN && i+j < (*sp_vWords).size(); ++j)
				{
					cache_it it = cache->find((*sp_vWords)[i+j]);
					if (TermNotCached(it))
					{
						cWords.push_back((*sp_vWords)[i+j]);
						if (!wInList.empty())
							wInList.append(L",");

						wInList.append(L"'");
						wInList.append(EscapeSQLQueryValue((*sp_vWords)[i+j]));
						wInList.append(L"'");
					}
				}

				// cache uncached tracked terms
				if (!wInList.empty())
				{
					wstring wQuery;
					wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from ");
					wQuery.append(wTableSetPrefix);
					wQuery.append(L"words where WoLgID = ");
					wQuery.append(wLgID);
					wQuery.append(L" AND WoTextLC in (");
					wQuery.append(wInList);
					wQuery.append(L")");

					EnterCriticalSection(&CS_UseDBConn);
					DoExecuteDirect(_T("1612isjdlfij"), wQuery);
					while (tStmt.fetch_next())
					{
						wstring wLC = tStmt.field(1).as_string();
						TermRecord rec(tStmt.field(2).as_string());
						rec.wRomanization = tStmt.field(3).as_string();
						rec.wTranslation = tStmt.field(4).as_string();
						cache->insert(unordered_map<wstring,TermRecord>::value_type(wLC, rec));
						cWords.remove(wLC);
					}
					tStmt.free_results();
					LeaveCriticalSection(&CS_UseDBConn);
				}
				// at this point all terms in list that are tracked are cached

				for (wstring& word : cWords)
					cache->insert(unordered_map<wstring,TermRecord>::value_type(word, TermRecord(L"0")));

				if (i == 0) // we're in the first loop, having cahced the first block of words
				{
					Helper_Notify_Thread_CachePageWords(pbHeadStartComplete, p_cv);
				}

				if (bShuttingDown)
					return;
			}
		}
	}
	void GetPageWords(IHTMLDocument2* pDoc, vector<wstring>& vWords)
	{
		if (!pDoc)
			return;

		IDOMTreeWalker* pDocTree = GetDocumentTree(pDoc); SmartCOMRelease scDocTree(pDocTree);
		if (!pDocTree)
			return;

		IDispatch* pNextNodeDisp = nullptr;
		IHTMLDOMNode* pNextNode = nullptr;
		IHTMLElement* pElement = nullptr;
		IHTMLDOMTextNode* pText = nullptr;
		bool bAwaitElement = false;
		HRESULT hr;

		// loop regex setup
		wstring regPtn = L"[" + WordChars + L"]";
		if (bWithSpaces)
			regPtn.append(L"+");
		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator rend;
		
		// this while loop manages reference counts manually
		while(SUCCEEDED(pDocTree->nextNode(&pNextNodeDisp)) && pNextNodeDisp)
		{
			long lngType;
			pNextNode = chaj::COM::GetAlternateInterface<IDispatch,IHTMLDOMNode>(pNextNodeDisp);
			if (pNextNode)
			{
				hr = pNextNode->get_nodeType(&lngType);
				if (SUCCEEDED(hr))
				{
					if (lngType == chaj::DOM::NODE_TYPE_ELEMENT)
					{
						bAwaitElement = false;
						pElement = chaj::COM::GetAlternateInterface<IHTMLDOMNode,IHTMLElement>(pNextNode);
						if (pElement)
						{
							wstring wTag = chaj::DOM::GetTagFromElement(pElement);
							if (wcsicmp(wTag.c_str(), L"script") == 0 || \
								wcsicmp(wTag.c_str(), L"style") == 0 || \
								wcsicmp(wTag.c_str(), L"textarea") == 0 || \
								wcsicmp(wTag.c_str(), L"iframe") == 0 || \
								wcsicmp(wTag.c_str(), L"noframe") == 0)
							{
								bAwaitElement = true;
							}
						}
						pElement->Release();
					}
					else if (!bAwaitElement && lngType == chaj::DOM::NODE_TYPE_TEXT)
					{
						pText = chaj::COM::GetAlternateInterface<IHTMLDOMNode,IHTMLDOMTextNode>(pNextNode);
						if (pText)
						{
							_bstr_t bText;
							wstring wText;
							pText->get_data(bText.GetAddress());
							if (bText.length())
								wText = bText;

							if (!wText.empty())
							{
								wregex_iterator regit(wText.begin(), wText.end(), wrgx);
								while (regit != rend)
								{
									vWords.push_back(regit->str());
									++regit;
								}
							}
						}
						pText->Release();
					}
				}
				pNextNode->Release();
			}
			pNextNodeDisp->Release();
			if (bShuttingDown)
				break;
		}
	}
	IWebBrowser2* UnstreamBrowser_Helper__Thread_OnPageFullyLoaded(LPSTREAM pBrowserStream)
	{
		IWebBrowser2* pBrowser = nullptr;
		HRESULT hr = CoGetInterfaceAndReleaseStream(pBrowserStream, IID_IWebBrowser2, reinterpret_cast<LPVOID*>(&pBrowser));
		if (FAILED(hr) || !pBrowser)
			return nullptr;
		else
			return pBrowser;
	}
	bool StartWordCache_Helper__Thread_OnPageFullyLoaded(IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords, condition_variable& cv, bool& bHeadStartComplete)
	{
		LPSTREAM pDocStream = nullptr;
		HRESULT hr = CoMarshalInterThreadInterfaceInStream(IID_IHTMLDocument2, pDoc, &pDocStream);
		if (SUCCEEDED(hr) && pDocStream)
		{
			cpThreads.push_back(new std::thread(&LwtBho::Thread_CachePageWords, this, pDocStream, sp_vWords, &cv, &bHeadStartComplete));
			return true;
		}
		else
			return false;
	}
	bool StartMWCache_Helper__Thread_OnPageFullyLoaded(IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords)
	{
		LPSTREAM pDocStream = nullptr;
		HRESULT hr = CoMarshalInterThreadInterfaceInStream(IID_IHTMLDocument2, pDoc, &pDocStream);
		if (SUCCEEDED(hr) && pDocStream)
		{
			cpThreads.push_back(new std::thread(&LwtBho::Thread_CachePageMWTerms, this, pDocStream, sp_vWords));
			return true;
		}
		else
			return false;
	}
	bool StartPageHighlightThread_Helper__Thread_OnPageFullyLoaded(IHTMLDocument2* pDoc)
	{
		LPSTREAM pDocStream = nullptr;
		HRESULT hr = CoMarshalInterThreadInterfaceInStream(IID_IHTMLDocument2, pDoc, &pDocStream);
		if (SUCCEEDED(hr) && pDocStream)
		{
			cpThreads.push_back(new std::thread(&LwtBho::Thread_HighlightPage, this, pDocStream));
			return true;
		}
		else
			return false;
	}
	void Thread_OnPageFullyLoaded(LPSTREAM pBrowserStream)
	{
		if (!pBrowserStream)
			return;

		IWebBrowser2* pBrowser = UnstreamBrowser_Helper__Thread_OnPageFullyLoaded(pBrowserStream); SmartCOMRelease scBrowser(pBrowser, true); // AddRef and schedule Release
		if (!pBrowser)
			return;

		IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser); SmartCOMRelease scDoc(pDoc);
		if (!pDoc)
			return;

		//wstring wBody;
		//IHTMLElement* pBody = GetBodyFromDoc(pDoc);
		//GetBodyContent(wBody, pBody, pDoc);
		AppendCss(pDoc);
		AppendHtmlBlocks(pDoc);
		AppendJavascript(pDoc);

		IHTMLElement* pSetupLink = chaj::DOM::GetElementFromId(L"lwtSetupLink", pDoc); SmartCOMRelease scSetupLink(pSetupLink);
		assert(pSetupLink);
		dhr(pSetupLink->click());

		// mutex and conditional variable to signal a batch of cached words
		std::mutex m;
		std::condition_variable cv;
		bool bHeadStartComplete = false;

		vector<wstring> vWords;
		GetPageWords(pDoc, vWords);

		shared_ptr<vector<wstring>> sp_vWords = make_shared<vector<wstring>>(vWords);

		// start threads
		StartMWCache_Helper__Thread_OnPageFullyLoaded(pDoc, sp_vWords);
		StartWordCache_Helper__Thread_OnPageFullyLoaded(pDoc, sp_vWords, cv, bHeadStartComplete);
		
		while (!bHeadStartComplete)
		{
			std:mutex n;
			std::unique_lock<mutex> lk(n);
			cv.wait(lk);
			if (bShuttingDown)
				return;
		}

		// start thread after above notification
		StartPageHighlightThread_Helper__Thread_OnPageFullyLoaded(pDoc);
	}
	void Thread_HighlightPage(LPSTREAM pDocStream)
	{
		if (!pDocStream)
			return;

		unordered_set<wstring> usetPageTerms; // used to track which terms have had an html hidden div record appended

		IHTMLDocument2* pDoc = nullptr;
		HRESULT hr = CoGetInterfaceAndReleaseStream(pDocStream, IID_IHTMLDocument2, reinterpret_cast<LPVOID*>(&pDoc));
		SmartCOMRelease scDoc(pDoc, true); // AddRef and schedule Release
		if (FAILED(hr) || !pDoc)
			return;

		IDOMTreeWalker* pDocTree = GetDocumentTree(pDoc); SmartCOMRelease scDocTree(pDocTree);
		if (!pDocTree)
			return;

		IDispatch* pNextNodeDisp = nullptr;
		IDispatch* pAfterTextDisp = nullptr;
		IHTMLDOMNode* pNextNode = nullptr;
		IHTMLElement* pElement = nullptr;
		IHTMLDOMTextNode* pText = nullptr;
		bool bAwaitElement = false;
		unsigned int count = 0;
		
		// this while loop manages reference counts manually
		while(SUCCEEDED(pDocTree->nextNode(&pNextNodeDisp)) && pNextNodeDisp)
		{
			long lngType;
			pNextNode = chaj::COM::GetAlternateInterface<IDispatch,IHTMLDOMNode>(pNextNodeDisp);
			if (pNextNode)
			{
				hr = pNextNode->get_nodeType(&lngType);
				if (SUCCEEDED(hr))
				{
					if (lngType == chaj::DOM::NODE_TYPE_ELEMENT)
					{
						bAwaitElement = false;
						pElement = chaj::COM::GetAlternateInterface<IHTMLDOMNode,IHTMLElement>(pNextNode);
						if (pElement)
						{
							wstring wTag = chaj::DOM::GetTagFromElement(pElement);
							if (wcsicmp(wTag.c_str(), L"span") == 0)
							{
								if (!GetAttributeValue(pElement, L"lwtTerm").empty())
									bAwaitElement = true;
							}
							else if (wcsicmp(wTag.c_str(), L"script") == 0 || \
								wcsicmp(wTag.c_str(), L"style") == 0 || \
								wcsicmp(wTag.c_str(), L"textarea") == 0 || \
								wcsicmp(wTag.c_str(), L"iframe") == 0 || \
								wcsicmp(wTag.c_str(), L"noframe") == 0)
							{
								bAwaitElement = true;
							}
						}
						pElement->Release();
					}
					else if (!bAwaitElement && lngType == chaj::DOM::NODE_TYPE_TEXT)
					{
						dhr(pDocTree->nextNode(&pAfterTextDisp)); // grab the node that follows this text node as a bookmark
						pText = chaj::COM::GetAlternateInterface<IHTMLDOMNode,IHTMLDOMTextNode>(pNextNode);
						if (pText)
							ParseTextNode_AllAtOnce(pText, pDoc, usetPageTerms);
						pText->Release();
						if (pAfterTextDisp) // kinda hackish, but take whatever node now precedes the bookmark so loop will increment back to bookmark
						{
							pAfterTextDisp->Release();
							hr = pDocTree->previousNode(&pAfterTextDisp);
							if (SUCCEEDED(hr) && pAfterTextDisp)
								pAfterTextDisp->Release();
						}
					}
				}
			}
			pNextNode->Release();
			pNextNodeDisp->Release();
			if (bShuttingDown)
				break;
		}
	}
	static void DeleteHandles()
	{
		for (vector<HANDLE>::size_type i = 0; i < vDelete.size(); ++i)
			DeleteObject(vDelete[i]);
		vDelete.clear();
	}
	static INT_PTR CALLBACK DlgProc_CtrlDlg(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		switch(uMsg)
		{
			case WM_COMMAND:
			{
				switch (HIWORD(wParam))
				{
					case BN_CLICKED:
					{
						switch (LOWORD(wParam))
						{
							case IDC_BTN_FORCELOAD:
							{
								EndDialog(hDlg, CTRL_FORCE_PARSE);
								break;
							}
							case IDC_BTN_LANGCHANGE:
							{
								EndDialog(hDlg, CTRL_CHANGE_LANG);
								break;
							}
						}
					}
				}
			}
		}

		return static_cast<INT_PTR>(0);
	}
	//static INT_PTR CALLBACK DlgProc_ChangeStatus(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	//{
	//	switch(uMsg)
	//	{
	//		case WM_INITDIALOG:
	//		{
	//			HBITMAP hBmp;
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS1));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON1),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS2));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON2),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS3));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON3),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS4));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON4),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS5));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON5),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS0));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON6),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS99));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON7),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS98));
	//			vDelete.push_back(hBmp);
	//			SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON8),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
	//			
	//			if (lParam)
	//			{
	//				MainDlgStruct* pMDS = reinterpret_cast<MainDlgStruct*>(lParam);
	//				vector<wstring>& vTerms = pMDS->Terms;
	//				if (vTerms.size() > 0)
	//				{
	//					for (int i = 0; i < vTerms.size(); ++i)
	//					{
	//						SendMessageW(GetDlgItem(hwndDlg, IDC_COMBO1), CB_ADDSTRING, NULL, reinterpret_cast<LPARAM>(vTerms[i].c_str()));
	//					}

	//					SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_SETCURSEL, (WPARAM)0, (LPARAM)0);
	//				}
	//			}
	//			return (INT_PTR) TRUE;
	//		}
	//		case WM_COMMAND:
	//		{
	//			switch(HIWORD(wParam))
	//			{
	//				case BN_CLICKED:
	//				{
	//					DlgResult* pDR;

	//					int wp = LOWORD(wParam);
	//					if (wp == IDC_STATIC_LINK)
	//					{
	//						DeleteHandles();
	//						pDR = new DlgResult;
	//						pDR->nStatus = 100;
	//						EndDialog(hwndDlg, (INT_PTR)pDR);
	//						break;
	//					}
	//					if (wp == IDC_BUTTON1 || wp == IDC_BUTTON2 || wp == IDC_BUTTON3 || wp == IDC_BUTTON4 || wp == IDC_BUTTON5 || wp == IDC_BUTTON6 || wp == IDC_BUTTON7 || wp == IDC_BUTTON8)
	//					{
	//						wchar_t wszEntry[LWT_MAXWORDLEN+1];
	//						int nCurSel = (int)SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_GETCURSEL, 0, 0);
	//						SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_GETLBTEXT, (WPARAM)nCurSel, reinterpret_cast<LPARAM>(wszEntry));
	//						pDR = new DlgResult;
	//						pDR->wstrTerm = wszEntry;		
	//					}

	//					switch(LOWORD(wParam))
	//					{
	//						case IDC_BUTTON1:
	//						{
	//							
	//							DeleteHandles();
	//							pDR->nStatus = 1;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON2:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 2;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON3:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 3;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON4:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 4;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON5:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 5;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON6:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 0;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON7:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 99;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//						case IDC_BUTTON8:
	//						{
	//							DeleteHandles();
	//							pDR->nStatus = 98;
	//							EndDialog(hwndDlg, (INT_PTR)pDR);
	//							break;
	//						}
	//					}
	//				}
	//				case CBN_SELCHANGE:
	//				{
	//					wchar_t wszEntry[LWT_MAXWORDLEN+1];
	//					int nCurSel = (int)SendMessage(reinterpret_cast<HWND>(lParam), CB_GETCURSEL, 0, 0);
	//					SendMessage(reinterpret_cast<HWND>(lParam), CB_GETLBTEXT, (WPARAM)nCurSel, reinterpret_cast<LPARAM>(wszEntry));
	//					SetWindowText(GetDlgItem(hwndDlg, IDC_EDIT1), wszEntry);
	//					break;
	//				}
	//			}
	//			return 0;
	//		}
	//		case WM_CLOSE:
	//		{
	//			DeleteHandles();
	//			DlgResult* pDR = new DlgResult;
	//			pDR->nStatus = 100;
	//			EndDialog(hwndDlg, (INT_PTR)pDR);
	//		}
	//	}
	//	return (INT_PTR) FALSE;
	//}
	bool IsMultiwordTerm(const wstring& wTerm)
	{
		if (bWithSpaces == false && wTerm.size() > 1)
			return true;
		if (bWithSpaces == true && wTerm.find(L" ") != wstring::npos)
			return true;
		return false;
	}
	void EnsureIterator(IHTMLDocument2* pDoc)
	{
		if (pNI)
			pNI->Release();

		IHTMLElement* pBody = GetBodyFromDoc(pDoc);
		DOMIteratorFilter filter(&FilterNodes_LWTTerm);
		pNI = GetNodeIteratorWithFilter(pDoc, pBody, dynamic_cast<IDispatch*>(&filter));
		pBody->Release();
	}
	void GetDropdownPrefixList(vector<wstring>& vPrefixes)
	{
		vPrefixes.clear();
		wstring wQuery(L"show tables like '%_settings'");
		if (tStmt.execute_direct(tConn, wQuery))
		{
			while(tStmt.fetch_next())
			{
				wstring wTableName = tStmt.field(1).as_string();
				vPrefixes.push_back(wstring(wTableName, 0, wTableName.size() - 9));
			}

			tStmt.free_results();
		}
	}
	void GetDropdownTermList(vector<wstring>& Terms, IHTMLElement* element)
	{
		IDispatch* pElemDisp = GetAlternateInterface<IHTMLElement,IDispatch>(element);
		if (!pElemDisp) return;

		wstring wAggTerm = GetAttributeValue(element, L"lwtTerm");
		if (wAggTerm.size() <= 0)
			return;

		Terms.push_back(wAggTerm);
		
		pTW->putref_currentNode(pElemDisp);
		int nCount = 1;
		IDispatch* pCur = NULL;
		IHTMLElement* pEl = NULL;
		wstring wClass;
		while (nCount < 9)
		{
			pTW->nextNode(&pCur);
			if (pCur)
			{
				pEl = GetAlternateInterface<IDispatch,IHTMLElement>(pCur);
				wClass = GetAttributeValue(pEl, L"class");
				_bstr_t bstrStat;
				pEl->get_className(bstrStat.GetAddress());
				wClass = bstrStat;
				if (wClass == L"lwtSent")
					break;
				wstring wTerm = GetAttributeValue(pEl, L"lwtTerm");
				if (wTerm.size() > 0 && !IsMultiwordTerm(wTerm))
				{
					if (bWithSpaces)
						wAggTerm.append(L" ");
					wAggTerm.append(wTerm);
					Terms.push_back(wAggTerm);
					++nCount;
				}
				pEl->Release();
			}
		}
	}
	bool ElementPartOfLink(IHTMLElement* pElement)
	{
		IHTMLAnchorElement* pAnchor = nullptr;

		do
		{
			pAnchor = GetAlternateInterface<IHTMLElement,IHTMLAnchorElement>(pElement);
			if (pAnchor) return true;

			pElement->get_parentElement(&pElement);

		} while (pElement);

		return false;
	}
	void HandleRightClick()
	{
		IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc) return;

		IHTMLEventObj* pEvent = GetEventFromDocument(pDoc);
		if (!pEvent)
		{
			pDoc->Release();
			return;
		}

		VARIANT_BOOL vBool;
		HRESULT hr = pEvent->get_ctrlKey(&vBool);
		if (FAILED(hr))
		{
			pEvent->Release();
			pDoc->Release();
			TRACE(L"Leaving HandleRightClick with error: could not retrieve Ctrl Key state.\n");
			return;
		}

		if (vBool == VARIANT_TRUE)
		{
			SetEventReturnFalse(pEvent);
			INT_PTR res = DialogBox(hInstance, MAKEINTRESOURCE(IDD_CTRL_DIALOG), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_CtrlDlg));
			if (res == CTRL_FORCE_PARSE)
				ForcePageParse(pDoc);
			else if (res == CTRL_CHANGE_LANG)
				HandleChangeLang();
		}
	}
	void HandleChangeLang()
	{
		// show dialog
		// get lang choice
		// set lwt db choice
		// call on lang change -- which should flush cache
		// reparse page if on lang change doesn't
	}
	void ForcePageParse(IHTMLDocument2* pDoc)
	{
		ParsePage(pDoc);
	}
	HRESULT SetEventReturnFalse(IHTMLEventObj* pEvent)
	{
		VARIANT varRetVal; VariantInit(&varRetVal); varRetVal.vt = VT_BOOL; varRetVal.boolVal = VARIANT_FALSE;
		HRESULT hr = pEvent->put_returnValue(varRetVal);
		VariantClear(&varRetVal);
		if (FAILED(hr))
			TRACE(L"Encountered unexpected hresult in SetEventReturnFalse.\n");
		return hr;
	}
	inline bool TermInCache(const wstring& wTerm)
	{
		return cache->find(wTerm) != cache->end();
		//unordered_map<wstring,TermRecord>::iterator it = cache->find(wTerm);
		//return it != cache->end();
	}
	void oldPopupDialog()
	{
		//else
		//{

		//	wstring wTerm = GetAttributeValue(pElement, L"lwtTerm");
		//	if (wTerm.size() <= 0)
		//		return;

		//	//_bstr_t bstrStat;
		//	BSTR bstrStat;
		//	HRESULT hr = pElement->get_className(&bstrStat);//bstrStat.GetAddress());
		//	if (bstrStat == NULL)
		//		return;

		//	wstring wOldStatus(bstrStat);
		//	SysFreeString(bstrStat);
		//	wstring wOldStatNum(wstring(wOldStatus, wStatIntro.size()));
		//	bool isAddOp = false;
		//	if (wOldStatNum == L"0")
		//		isAddOp = true;


		//	MainDlgStruct mds(ElementPartOfLink(pElement));
		//	DlgResult* pDR = NULL;
		//	int res;
		//	if (IsMultiwordTerm(wTerm))
		//	{
		//		if (mds.bOnLink)
		//			pDR = (DlgResult*)DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG2), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus));
		//		else
		//			pDR = (DlgResult*)DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG1), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus));
		//		if ((int)pDR == -1)
		//		{
		//			mb("Unable to render dialog box, or other related error.", "2232idhf");
		//			return;
		//		}

		//		res = pDR->nStatus;
		//		if (res != 100)
		//		{
		//			assert(res == 0 || res == 1 || res == 2 || res == 3 || res == 4 || res == 5 || res == 98 || res == 99);
		//			ChangeTermStatus(wTerm, to_wstring(res), pDoc);
		//			if (mds.bOnLink)
		//				SetEventReturnFalse(pEvent);
		//		}
		//	}
		//	else
		//	{
		//		GetDropdownTermList(mds.Terms, pElement);
		//		if (mds.bOnLink)
		//			pDR = (DlgResult*)DialogBoxParam(hInstance, MAKEINTRESOURCE(IDD_DIALOG12), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus), (LPARAM)&mds);
		//		else
		//			pDR = (DlgResult*)DialogBoxParam(hInstance, MAKEINTRESOURCE(IDD_DIALOG11), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus), (LPARAM)&mds);

		//		if ((int)pDR == -1)
		//		{
		//			mb("Unable to render dialog box, or other related error.", "2232idhf");
		//			return;
		//		}

		//		res = pDR->nStatus;
		//		if (res != 100)
		//		{
		//			assert(res == 0 || res == 1 || res == 2 || res == 3 || res == 4 || res == 5 || res == 98 || res == 99);
		//			if (IsMultiwordTerm(pDR->wstrTerm))
		//			{
		//				if (!TermInCache(pDR->wstrTerm))
		//					AddNewTerm(pDR->wstrTerm, to_wstring(res), pDoc);
		//				else
		//					mb("Term is already defined. Isn't it?");
		//			}
		//			else
		//			{
		//				if (isAddOp == true)
		//					AddNewTerm(wTerm, to_wstring(res), pDoc);
		//				else
		//					ChangeTermStatus(wTerm, to_wstring(res), pDoc);
		//			}

		//			VARIANT varRetVal; VariantInit(&varRetVal); varRetVal.vt = VT_BOOL; varRetVal.boolVal = VARIANT_FALSE;
		//			pEvent->put_returnValue(varRetVal);
		//			VariantClear(&varRetVal);
		//		}
		//	}

		//	pElement->Release();
		//	pEvent->Release();
		//	pDoc->Release();
		//	delete pDR;
		//}
	}
	wstring GetInfoTrans(IHTMLDocument2* pDoc)
	{
		IHTMLElement* pNewTrans = GetElementFromId(L"lwtshowtrans", pDoc);
		assert(pNewTrans);

		_bstr_t bNewTrans;
		dhr(pNewTrans->get_innerText(bNewTrans.GetAddress()));
		wstring wNewTrans(bNewTrans);

		pNewTrans->Release();

		wstring_replaceAll(wNewTrans, L"\r", L"");
		return wNewTrans;
	}
	wstring GetInfoRom(IHTMLDocument2* pDoc)
	{
		IHTMLElement* pNewRom = GetElementFromId(L"lwtshowrom", pDoc);
		assert(pNewRom);

		_bstr_t bNewRom;
		dhr(pNewRom->get_innerText(bNewRom.GetAddress()));
		wstring wNewRom(bNewRom);

		pNewRom->Release();

		wstring_replaceAll(wNewRom, L"\r", L"");
		return wNewRom;
	}
	/*
		TermNotTracked does not search the cache; it reports only on the state of the current iterator.

		There are two expected cases that must be covered:
		1) a multiword term that is not tracked and so it is not in the cache at all
		2) a single word term that is not tracked, but being on the page, is in the cache with a 0 status

		The result is logaically equivalent to "not in cache, or, there but with a 0 status" which should cover the
		full intent of the function's name.
	*/
	bool TermNotTracked(cache_it it)
	{
		return (it == cache->end() || it->second.wStatus == L"0"); // an untracked MW term || an untracked word
	}
	bool TermNotTracked(cache_cit it)
	{
		return (it == cache->end() || it->second.wStatus == L"0"); // an untracked MW term || an untracked word
	}
	/*
		TermNotCached does not search the cache; it reports only on the state of the current iterator.
	*/
	bool TermNotCached(cache_it it)
	{
		return (it == cache->end());
	}
	bool TermNotCached(cache_cit it)
	{
		return (it == cache->end());
	}
	/*
		TermNotCached does not search the cache; it reports only on the state of the current iterator.
	*/
	bool TermCached(cache_it it)
	{
		return (it != cache->cend());
	}
	bool TermCached(cache_cit it)
	{
		return (it != cache->cend());
	}

	/*
		UpdateTrackedTerm

		Returns true is update is successful, or false otherwise.
	*/
	bool UpdateTrackedTerm(const wstring& wCurInfoTerm, const wstring& wNewTrans, const wstring& wNewRom, cache_it& it)
	{
		bool bRetVal = false;

		wstring wQuery = L"update ";
		wQuery += wTableSetPrefix;
		wQuery.append(L"words set WoTranslation = '");
		wQuery += this->EscapeSQLQueryValue(wNewTrans);
		wQuery.append(L"', WoRomanization = '");
		wQuery += this->EscapeSQLQueryValue(wNewRom);
		wQuery.append(L"' where WoTextLC = '");
		wQuery += wCurInfoTerm;
		wQuery.append(L"' and WoLgID = ");
		wQuery += this->wLgID;
				
		if (DoExecuteDirect(L"4067xhidjf", wQuery))
		{
			it->second.wTranslation = wNewTrans;
			it->second.wRomanization = wNewRom;
			bRetVal = true;
		}

		return bRetVal;
	}
	/*	UpdateWebPageDivRec (Update Web Page Term Record Div Element)

		Given a term's cache id by which it's record can be located, this function will update
		the contents of that resident div record
	*/
	bool UpdateWebPageDivRec(unsigned int uTermID, const wstring& wNewTrans, const wstring& wNewRom, IHTMLDocument2* pDoc)
	{
		chaj::util::CatchSentinel<HRESULT> cs(S_OK, true);

		wstring wCurTermId(L"lwt");
		wCurTermId.append(to_wstring(uTermID));

		IHTMLElement* pCurTermDivRec = GetElementFromId(wCurTermId, pDoc); SmartCOMRelease scCurTermDivRec(pCurTermDivRec);
		if (pCurTermDivRec)
		{
			cs += SetAttributeValue(pCurTermDivRec, L"lwttrans", wNewTrans);
			cs += SetAttributeValue(pCurTermDivRec, L"lwtrom", wNewRom);
		}

		if (cs.Seen())
			return false;
		else
			return true;
	}
	void HandleUpdateTermInfo(IHTMLDocument2* pDoc)
	{
		IHTMLElement* pCurInfoTerm = GetElementFromId(L"lwtcurinfoterm", pDoc); SmartCOMRelease scCurInfoTerm(pCurInfoTerm);
		if (!pCurInfoTerm) return;

		wstring wCurInfoTerm = GetAttributeValue(pCurInfoTerm, L"lwtterm");
		if (!wCurInfoTerm.size()) return;

		cache_it it = cache->find(wCurInfoTerm);

		if (TermNotTracked(it))
		{
			AddNewTerm(wCurInfoTerm, L"1", pDoc);
			if (TermNotCached(it)) // then update iterator to point to recent addition
				it = cache->find(wCurInfoTerm);
		}

		assert(it != cache->end()); // logically impossible
		if (it == cache->end()) // crash guard
			return;

		// update tracked term
		wstring wNewTrans = GetInfoTrans(pDoc);
		wstring wNewRom = GetInfoRom(pDoc);
		if (!UpdateTrackedTerm(wCurInfoTerm, wNewTrans, wNewRom, it))
			return;

		// update webpage
		UpdateWebPageDivRec(it->second.uIdent, wNewTrans, wNewRom, pDoc);
		SendLwtCommand(L"closeInfoEdit", pDoc);
	}
	void HandleStatChange(IHTMLElement* pElement, IHTMLDocument2* pDoc)
	{
		IHTMLElement* pLast = GetElementFromId(L"lwtlasthovered", pDoc); SmartCOMRelease scLast(pLast);
		if (!pLast) return;

		wstring wLast = GetAttributeValue(pLast, L"lwtterm");
		wstring wCurStat = GetAttributeValue(pLast, L"lwtstat");
		wstring wNewStat = GetAttributeValue(pElement, L"lwtstat");

		if (!wLast.size())
			return;

		if (wCurStat == L"0")
			AddNewTerm(wLast, wNewStat, pDoc);
		else
			ChangeTermStatus(wLast, wNewStat, pDoc);
			
		SetAttributeValue(pLast, L"lwtstat", wNewStat);
	}
	/*
		ClearJSCommand clears hidden html elements that pass commands and arguments

		This is an attempt to ensure clean transactions.
	*/
	void ClearJSCommand(IHTMLDocument2* pDoc)
	{
		if (!pDoc) return;

		IHTMLElement* pJSCommand = GetElementFromId(L"lwtJSCommand", pDoc); SmartCOMRelease scJSCommand(pJSCommand);
		if (!pJSCommand) return;

		chaj::DOM::SetAttributeValue(pJSCommand, L"lwtAction", L"");
		chaj::DOM::SetAttributeValue(pJSCommand, L"value", L"");
	}
	bool HandleFillMWTerm(IHTMLDocument2* pDoc)
	{
		if (!pDoc) return false;

		IHTMLElement* pCurInfoTerm = GetElementFromId(L"lwtcurinfoterm", pDoc); SmartCOMRelease scCurInfoTerm(pCurInfoTerm);
		if (!pCurInfoTerm) return false;

		wstring wCurInfoTerm = GetAttributeValue(pCurInfoTerm, L"lwtterm");
		if (!wCurInfoTerm.size()) return false;

		cache_cit it = cache->cfind(wCurInfoTerm);
		if (TermNotCached(it)) return false;

		IHTMLElement* pTrans = GetElementFromId(L"lwtshowtrans", pDoc); SmartCOMRelease scTrans(pTrans);
		IHTMLElement* pRom = GetElementFromId(L"lwtshowtrans", pDoc); SmartCOMRelease scRom(pRom);
		if (!pTrans || !pRom)
			return false;

		chaj::util::CatchSentinel<HRESULT> cs(S_OK, true);
		cs += chaj::DOM::SetElementInnerText(pTrans, it->second.wTranslation);
		cs += chaj::DOM::SetElementInnerText(pRom, it->second.wRomanization);

		if (cs.Seen())	return false;
		else			return true;
	}
	void ReloadWebpage(IHTMLDocument2* pDoc)
	{
		SendLwtCommand(L"reloadPage", pDoc);
	}
	void GetLgID()
	{
		wstring wQuery(L"select StValue from ");
		wQuery += wTableSetPrefix;
		wQuery.append(L"settings where StKey = 'currentlanguage'");

		if (!tStmt.execute_direct(tConn, wQuery))
		{
			mb(L"Could not retrieve lwt language ID from settings.");
			return;
		}
		else
		{
			if (tStmt.fetch_next() == false)
			{
				wLgID = L"";
				mb(L"No languages have been defined for this TableSet -- or at least one is not selected.");
			}
			else
				wLgID = tStmt.field(1).as_string();
			tStmt.free_results();
		}
	}
	void HandleClick()
	{
		IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser); SmartCOMRelease scDoc(pDoc);
		if (!pDoc) return;

		IHTMLEventObj* pEvent = GetEventFromDocument(pDoc); SmartCOMRelease scEvent(pEvent);
		if (!pEvent) return;

		IHTMLElement* pElement = GetClickedElementFromEvent(pEvent, pDoc); SmartCOMRelease scElement(pElement);
		if (!pElement) return;

		wstring wActionReq = GetAttributeValue(pElement, L"lwtAction");
		wstring wStatChange = GetAttributeValue(pElement, L"lwtstatchange");
		wstring wActionValue;
		if (wActionReq.size())
		{
			wActionValue = GetAttributeValue(pElement, L"value");
			ClearJSCommand(pDoc);
		}

		if (wStatChange.size())
			HandleStatChange(pElement, pDoc);
		else if (wActionReq == L"lwtUpdateTermInfo")
			HandleUpdateTermInfo(pDoc);
		else if (wActionReq == L"FillMWTerm")
			HandleFillMWTerm(pDoc);
		else if (wActionReq == L"changeLang")
		{
			wstring wNewLgID = GetAttributeValue(pElement, L"value");
			if (wNewLgID.size())
			{
				this->wLgID = wNewLgID;
				OnLanguageChange(pDoc);
			}
		}
		else if (wActionReq == L"changeTableSet")
		{
			wTableSetPrefix = wActionValue;
				if (wTableSetPrefix.size())
					wTableSetPrefix += L"_";
			OnTableSetChange(pDoc);
		}
	}
	IHTMLElement* GetHTMLElementFromDisp(IDispatch* pDispElement)
	{
		IHTMLElement* res = NULL;
		pDispElement->QueryInterface(IID_IHTMLElement, (void**)&res);
		return res;
	}
	/*
		GetDocumentFromBrowser
		Caller contract: caller must release the interface returned if non-null
		Input conditions: a valid pointer to extant IWebBrowser2 interface
		Return value: returns nullptr on error, or a valid pointer to an extant IHTMLDocument2 interface
	*/

	IHTMLEventObj* GetEventFromDocument(IHTMLDocument2* pDoc)
	{
		IHTMLWindow2* window = NULL;
		HRESULT hr = pDoc->get_parentWindow(&window);
		if (FAILED(hr))
		{
			mb(L"get_parentWindow failed", L"2834hduufh");
			pDoc->Release();
			TRACE(L"%s", L"Leaving GetEventFromDocument with error result\n");
			return nullptr;
		}

		IHTMLEventObj* event;
		hr = window->get_event(&event);
		window->Release();
		if (FAILED(hr))
		{
			mb(L"get_event failed", L"534shefhhfu");
			pDoc->Release();
			TRACE(L"%s", L"Leaving GetEventFromDocument with error result\n");
			return nullptr;
		}

		return event;
	}
	IHTMLElement* GetClickedElementFromEvent(IHTMLEventObj* pEvent, IHTMLDocument2* pDoc)
	{
		long x, y;
		pEvent->get_clientX(&x);
		pEvent->get_clientY(&y);

		IHTMLElement* pElement;
		HRESULT hr = pEvent->get_srcElement(&pElement);
		//HRESULT hr = pDoc->elementFromPoint(x, y, &pElement);
		if (FAILED(hr))
		{
			TRACE(L"%s", L"Leaving GetClickedElementFromEvent with error result\n");
			return nullptr;
		}

		return pElement;
	}
	void OnDocumentComplete(DISPPARAMS* pDispParams)
	{
		if (pDispParams->cArgs >= 2 && pDispParams->rgvarg[1].vt == VT_DISPATCH)
		{
			IDispatch* pCur = pDispParams->rgvarg[1].pdispVal;
			if (pCur == pDispBrowser)
			{
				IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser); SmartCOMRelease scDoc(pDoc);
				ReceiveEvents(pDoc);
				LPSTREAM pBrowserStream = nullptr;
				HRESULT hr = CoMarshalInterThreadInterfaceInStream(IID_IWebBrowser2, this->pBrowser, &pBrowserStream);
				if (SUCCEEDED(hr) && pBrowserStream)
					cpThreads.push_back(new std::thread(&LwtBho::Thread_OnPageFullyLoaded, this, pBrowserStream));
			}
		}
	}
	HRESULT _stdcall Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		if (dispidMember == DISPID_HTMLDOCUMENTEVENTS_ONCLICK)
			HandleClick();
		else if (dispidMember == DISPID_BEFORENAVIGATE2)
			int i = 0;
		else if (dispidMember == DISPID_HTMLDOCUMENTEVENTS_ONCONTEXTMENU)
			HandleRightClick();
		else if (dispidMember == DISPID_DOCUMENTCOMPLETE)
			OnDocumentComplete(pDispParams);
		return S_OK;
		//if (dispidMember == DISPID_DOWNLOADBEGIN)
		//if (dispidMember == DISPID_DOWNLOADCOMPLETE)
		//if (dispidMember == DISPID_NAVIGATECOMPLETE2)
		//if (dispidMember == DISPID_ONQUIT)
	}

/* THIS IS POTENTIAL CODE FOR AFTER DOCUMENT COMPLETE */
/* it would allow javascript to directly access bho, with native calls like window.lwtbho.docommand(); */

//IHTMLWindow2* pWindow = nullptr;
//IDispatchEx* pDispEx = nullptr;
//IDispatch* pDispLwtBho = nullptr;
//pDoc->get_parentWindow(&pWindow);
//pWindow->QueryInterface(IID_IDispatchEx, reinterpret_cast<void**>(&pDispEx));
//this->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&pDispLwtBho));
//_bstr_t bBhoName(L"LwtBho");
//DISPID dispid;
//pDispEx->GetDispID(bBhoName, fdexNameEnsure, &dispid);
//DISPPARAMS params;
//VARIANT vPut; VariantInit(&vPut); vPut.vt = VT_DISPATCH; vPut.pdispVal = pDispLwtBho;
//params.cArgs = 1;
//params.cNamedArgs = 0;
//params.rgdispidNamedArgs = NULL;
//params.rgvarg = &vPut;
//pDispEx->Invoke(dispid, IID_NULL, NULL, DISPATCH_PROPERTYPUTREF, &params, NULL, NULL, NULL);

private:
	void GetCurrentTableSetPrefix();
	void LoadJavascriptFile();

	/* member variables */
	unordered_set<wstring> usetCacheMWFragments;

	// interfaces
	IUnknown* pSite;
	IWebBrowser2* pBrowser;
	IConnectionPoint* pCP;
	IConnectionPoint* pHDEv;
	IDOMNodeIterator* pNI;
	IDOMTreeWalker* pTW;
	IDispatch* pDispBrowser;

	chaj::DOM::DOMIteratorFilter* pFilter;
	chaj::DOM::DOMIteratorFilter* pFullFilter;
	HWND hBrowser;
	DWORD dwCookie, Cookie_pHDev;
	BOOL bDocumentCompleted;
	long ref;
	mysqlpp::Connection conn;
	tiodbc::connection tConn;
	tiodbc::statement tStmt;
	wstring wTableSetPrefix;
	wstring wLgID;
	wstring WordChars;
	wstring SentDelims;
	wstring SentDelimExcepts;
	wstring wstrDict1;
	wstring wstrDict2;
	wstring wstrGoogleTrans;
	wstring wJavascript;
	wstring wCss;
	vector<wstring> chunks;
	TokenStruct tokens, tknSansBookmarks, tknCanonical;
	bool bWithSpaces;
	_bstr_t bstrUrl;
	LWTCache* cache;
	unordered_map<wstring, LWTCache*> cacheMap;
	vector<std::thread*> cpThreads;
	// wregex
	wregex_iterator rend;
	wregex wWordSansBookmarksWrgx;
	wregex TagNotTagWrgx;
	bool bShuttingDown;
	CRITICAL_SECTION CS_UseDBConn;
};

inline void LwtBho::GetCurrentTableSetPrefix()
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

inline void LwtBho::LoadJavascriptFile()
{
	HRSRC rc = ::FindResource(hInstance, MAKEINTRESOURCE(IDR_LWTJAVASCRIPT01),
		MAKEINTRESOURCE(JAVASCRIPT));
	HGLOBAL rcData = ::LoadResource(hInstance, rc);
	DWORD size = ::SizeofResource(hInstance, rc);
	wJavascript = to_wstring(static_cast<const char*>(::LockResource(rcData)));
}

inline LwtBho::LwtBho()
{
#ifdef _DEBUG
	mb("Here's your chance to attach debugger...");
#endif
	bShuttingDown = false;

	pSite = nullptr;
	pBrowser = nullptr;
	pCP = nullptr;
	pHDEv = nullptr;
	pNI = nullptr;
	pTW = nullptr;
	pDispBrowser = nullptr;
	cache = nullptr;

	hBrowser = NULL;
	dwCookie = NULL;
	Cookie_pHDev = NULL;
	bDocumentCompleted = false;
	ref = 0;
	wLgID = L"";
	WordChars = L"";
	SentDelims = L"";
	SentDelimExcepts = L"";
	bWithSpaces = true;
	bstrUrl = "";

	LoadJavascriptFile();
	assert(this->wJavascript.size());
	LoadCssFile();
	assert(this->wCss.size());

	if (!InitializeCriticalSectionAndSpinCount(&CS_UseDBConn, 0x00000400))
		TRACE(L"Cound not initialize critical section CS_UseDBConn in LwtBho. 5110suhdg\n"); // hack: this needs to be addressed, but I didn't want to throw an exception from the constructor; research this

	InitializeWregexes();

	pFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTerm);
	pFullFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTermAndSent);

	//if (!tConn.connect(L"LWT", L"root", L""))
	if (!tConn.lwtConnect(hBrowser))
	{
		//mb(_T("Cannot connect to the Data Source"));
		//mb(tConn.last_error());
		return;
	}

	GetCurrentTableSetPrefix();
	OnTableSetChange(nullptr, false);
}

inline LwtBho::~LwtBho()
{
	bShuttingDown = true;

	// wait for all threads to shut down
	// thread is not deleted as this was causing issues for unknown reasons (would need troubleshooting)
	// LwtBho destructor will free all objects soon enough anyway
	for (unsigned int i = 0; i < cpThreads.size(); ++i)
	{
		if (cpThreads[i]->joinable())
			cpThreads[i]->join();
	}

	// all other threads have exited, safe to delete critical sections
	DeleteCriticalSection(&CS_UseDBConn);

	if (pFilter != nullptr)
		delete pFilter;

	for (auto cacheEntry : cacheMap)
	{
		delete cacheEntry.second;
	}
}

#endif