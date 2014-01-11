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
using namespace mysqlpp;
using namespace std;
using namespace chaj;
using namespace chaj::DOM;
using namespace chaj::COM;
using chaj::str::to_wstring;

typedef regex_iterator<wstring::const_iterator, wchar_t,regex_traits<wchar_t>> wregex_iterator;
typedef vector<wstring>::const_iterator wstr_c_it;
typedef pair<wstring,TermRecord> TermEntry;
const INT_PTR CTRL_FORCE_PARSE = static_cast<INT_PTR>(101);
const INT_PTR CTRL_CHANGE_LANG = static_cast<INT_PTR>(102);
const int MWSpanAltroSize = 7;
const int MAX_FIELD_WIDTH = 255;
const wstring wstrNewline(L"&#13;&#10;");
const wstring wTermDivider(L"___LwtTD___");

#define LWT_MAXWORDLEN 255
#define LWT_MAX_MWTERM_LENGTH 9
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

class LwtBho: public IObjectWithSite, public IDispatch
{
	friend class LWTCache;

public:
	LwtBho();
	~LwtBho();

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
#pragma warning( push )
#pragma warning( disable : 4100 )
	HRESULT STDMETHODCALLTYPE GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR*  ppTInfo)
	{
		return NOERROR;
	}
	HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId)
	{
		return NOERROR;
	}
#pragma warning( pop )

private:	
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
				AddNewMWSpans(wTerm, pDoc);
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
	void AddNewMWSpans(const wstring& wTerm, IHTMLDocument2* pDoc)
	{
		if (!wTerm.size() || !pDoc)
			return;

		vector<wstring> vParts;

		if (bWithSpaces)
		{
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

		SmartCom<IHTMLElement> scBody = GetBodyFromDoc(pDoc);
		if (!scBody)
			return;

		AppendTermDivRec(scBody, wTerm, it->second);
		AppendMultiWordSpan(out, wTerm, &(it->second), to_wstring(uCount));

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		if (!pRange)
			return;

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

			SmartCom<IHTMLElement> pStartElem;

			for (vector<wstring>::size_type i = 0; i < vParts.size(); ++i)
			{
				_bstr_t bMWTermPart(vParts[i].c_str());
				pRange->findText(bMWTermPart.GetBSTR(), 0, 0, &vBool);
				assert(vBool == VARIANT_TRUE);

				IHTMLElement* pMWPartSpan = nullptr;
				hr = pRange->parentElement(&pMWPartSpan); // release manually
				if (FAILED(hr) || !pMWPartSpan)
					return;

				while (pMWPartSpan && !GetAttributeValue(pMWPartSpan, L"lwtterm").size())
				{
					IHTMLElement* pHigherParent = nullptr;
					hr = pMWPartSpan->get_parentElement(&pHigherParent); // release manually
					if (FAILED(hr) || !pHigherParent)
						return;
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

				hr = chaj::DOM::TxtRange_CollapseToEnd(pRange);
				assert(SUCCEEDED(hr));

				hr = chaj::DOM::TxtRange_RevertEnd(pRange);
				assert(SUCCEEDED(hr));
			}

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
	HRESULT RemoveDomNodeByElement(IHTMLElement* pElement)
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
					HRESULT hr = RemoveDomNodeByElement(pSup);
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
		if (wNewStatus == L"0" && IsMultiwordTerm(wTerm))
			cache->erase(it);
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
		SmartCom<IHTMLElementCollection> pSpans = GetAllFromDoc_ByTag(pDoc, L"span");
		if (!pSpans)
			return;

		long lLength = 0;
		HRESULT hr = pSpans->get_length(&lLength);
		if (FAILED(hr) || lLength == 0)
			return;

		SmartCom<IDispatch> pDispElem;
		SmartCom<IHTMLElement> pSpan;
		VARIANT vIndex; VariantInit(&vIndex); vIndex.vt = VT_I4; 
		VARIANT vSubIndex; VariantInit(&vSubIndex); vSubIndex.vt = VT_I4; vSubIndex.lVal = 0;
		VARIANT varNewClassValue; VariantInit(&varNewClassValue); varNewClassValue.vt = VT_BSTR;
		wstring wNewStat = L"lwtStat" + wNewStatus;
		for (long i = 0; i < lLength; ++i)
		{
			vIndex.lVal = i;
			hr = pSpans->item(vIndex, vSubIndex, (IDispatch**)pDispElem);
			if (FAILED(hr) || !pDispElem)
				continue;

			pSpan = GetElementFromDisp(pDispElem);
			if (!pSpan)
				continue;

			wstring wAttValue = GetAttributeValue(pSpan, L"lwtTerm");
			if (wAttValue == wTerm)
			{
				_bstr_t bstrClassVal(wNewStat.c_str());
				pSpan->put_className(bstrClassVal.GetBSTR());
			}
		}
	}
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
	int FindProperCloseTagPos(const wstring& in, unsigned int nOpenTagPos)
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
	bool TagIsSelfClosed(const wstring& wTag)
	{
		wstring::size_type pos = wTag.find_last_not_of(L"> ");
		if (wTag[pos] != '/')
			return false;
		if (wTag.find(L">", pos) == wstring::npos)
			return false;

		return true;
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
			unsigned int pos = 0;

			while (pos != wstring::npos)
			{
				pos = wMWTerm.find(L" ", pos);
				usetCacheMWFragments.insert(wstring(wMWTerm, 0, pos));
				if (pos != wstring::npos)
					++pos;
			}
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
	wstring EscapeSQLQueryValue(const wstring& wQuery)
	{
		wstring out(wQuery);
		EscapeSQLQueryValue(out);
		return out;
	}
	wstring& EscapeSQLQueryValue(wstring& wQuery)
	{
		unsigned int pos = wQuery.find_first_of(L"'\\");
		if (pos == wstring::npos)
			return wQuery;

		unsigned int pos2;
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
	void ReceiveEvents(IHTMLDocument2* pDoc)
	{
		if (!pDoc)
			return;
		
		SmartCom<IConnectionPointContainer> aCPC;
		HRESULT hr = pDoc->QueryInterface(IID_IConnectionPointContainer, (void**)aCPC);
		if (FAILED(hr) || !aCPC)
			return;

		SmartCom<IConnectionPoint> pHDEv2;
		hr = aCPC->FindConnectionPoint(DIID_HTMLDocumentEvents, pHDEv2);
		if (FAILED(hr) || !pHDEv2)
			return;

		if (pHDEv != pHDEv2)
		{
			if (pHDEv != NULL)
			{
				pHDEv->Unadvise(Cookie_pHDev);
				pHDEv->Release();
			}
		
			pHDEv = pHDEv2.Relinquish();
			HRESULT hr = pHDEv->Advise(reinterpret_cast<IDispatch*>(this), &Cookie_pHDev);
		}
		else
			pHDEv2->Release();
	}
	void AppendSoloWordSpan(wstring& out, const wstring& wVisibleForm, const wstring& wCanonicalTerm, const TermRecord* pRec)
	{
		AppendTermSpan(out, wVisibleForm, wCanonicalTerm, pRec->wStatus, pRec->uIdent);
	}
	void AppendMultiWordSpan(wstring& out, const wstring& wCanonicalTerm, const TermRecord* pRec, const wstring& wWordCount)
	{
		out.append(L"<sup>");
		AppendTermSpan(out, wWordCount, wCanonicalTerm, pRec->wStatus, pRec->uIdent);
		out.append(L" </sup>");
	}
	void AppendTermSpan(wstring& out, const wstring& wVisibleForm, const wstring& wCanonicalTerm,  const wstring& wStatus, unsigned int Id)
	{
		out.append(L"<span class=\"lwtStat");
		out += wStatus;
		out.append(L"\" style=\"display:inline;margin:0px;color:black;\" lwtTerm=\"");
		out.append(wCanonicalTerm);
		out.append(L"\" onmouseover=\"");

		AppendTermSpanMouseoverAttribute(out, Id);

		out.append(L"\" onmouseout=\"lwtmout(event);\">");
		out.append(wVisibleForm);
		out.append(L"</span>");
	}
	void AppendTermSpanMouseoverAttribute(wstring& out, unsigned int Id)
	{
		out.append(L"lwtmover('lwt");
		out.append(to_wstring(Id));
		out.append(L"', event, this);");
	}
	void AppendJavascript(IHTMLDocument2* pDoc)
	{
		if (!pDoc)
		{
			TRACE(L"Argment error in AppendJavascript. 2513ishdllijj\n");
			return;
		}

		SmartCom<IHTMLElement> pHead = chaj::DOM::GetHeadFromDoc(pDoc);
		SmartCom<IHTMLElement> pBody = chaj::DOM::GetBodyFromDoc(pDoc);
		SmartCom<IHTMLElement> pScript = CreateElement(pDoc, L"script");
		SmartCom<IHTMLDOMNode> pScriptNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pScript);
		if (!pHead || !pBody || !pScript || !pScriptNode)
		{
			TRACE(L"Unable to append lwt javascript properly.\n");
			return;
		}

		SetAttributeValue(pScript, L"type", L"text/javascript");
		SetElementInnerText(pScript, wJavascript);

		bool DoHead = false;
		if (DoHead)
		{
			SmartCom<IHTMLDOMNode> pHeadNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pHead);
			SmartCom<IHTMLDOMNode> pNewScriptNode;
			HRESULT hr = pHeadNode->appendChild(pScriptNode, pNewScriptNode);
			if (FAILED(hr))
				TRACE(L"Unable to append javacript child. 23328ijf\n");
		}
		else
		{
			SmartCom<IHTMLDOMNode> pBodyNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pBody);
			SmartCom<IHTMLDOMNode> pNewScriptNode;
			HRESULT hr = pBodyNode->appendChild(pScriptNode, pNewScriptNode);
			if (FAILED(hr))
				TRACE(L"Unable to append javacript child. 2546suhfuugs\n");
		}
	}
	void wstring_replaceAll(wstring& out, const wstring& find, const wstring& replace)
	{
		unsigned int pos = 0;
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

		VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = s_pFullFilter;
		pDT->createTreeWalker(pRoot, SHOW_ELEMENT, &varFilter, VARIANT_TRUE, &pTW);
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

		IHTMLElement* pBody = GetBodyFromDoc(pDoc);
		if (!pBody)
			return;

		wstring out;
		out.append(
		L"<div id=\"lwtinlinestat\" style=\"display:none;position:absolute;\" onmouseout=\"lwtdivmout(event);\">"
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
	void AppendCss(IHTMLDocument2* pDoc)
	{
		HRESULT hr = AppendStylesToDoc(this->wCss, pDoc);
		assert(SUCCEEDED(hr));
	}
	void dhr(HRESULT hr) //(D)ebug (HR)esult
	{
		assert(SUCCEEDED(hr));
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
		if (FAILED(hr) || !pBrowser)
		{
			mb(_T("IWebBrowser2 fail"), _T("1495hfhoksj"));
			pSite->Release();
			pSite = NULL;
			return E_UNEXPECTED;
		}

		pDispBrowser = GetAlternateInterface<IWebBrowser2, IDispatch>(pBrowser);

		IConnectionPointContainer* pCPC;
		hr = pUnkSite->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
		if (FAILED(hr) || !pCPC)
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
		if (FAILED(hr) || !pCP)
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
			if (FAILED(hr) || !iThis)
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

		SmartCom<IHTMLElement> pRoot;
		pDoc->get_body(pRoot);
		if (!pRoot)
			return nullptr;

		SmartCom<IDocumentTraversal> pDT = GetAlternateInterface<IHTMLDocument2,IDocumentTraversal>(pDoc);
		if (!pDT)
			return nullptr;

		IDOMTreeWalker* pTW = nullptr;
		VARIANT varNull; varNull.vt = VT_NULL;
		hr = pDT->createTreeWalker(pRoot, SHOW_ELEMENT | SHOW_TEXT, &varNull, VARIANT_TRUE, &pTW);
		if (FAILED(hr) || !pTW)
			return nullptr;
		else
			return pTW;
	}
	cache_it CacheTermAndReturnIt(const wstring& wWord) // optimizable
	{
		list<wstring> cWord;
		cWord.push_back(wWord);
		EnsureRecordEntryForEachWord(cWord);
		return cache->find(wWord);
	}
	void MineTextNodeForWords_InsertPlaceHolders(IHTMLDOMTextNode* pText, IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords)
	{
		if (!pText || !pDoc)
			return;

		wstring wText = GetTextNodeText(pText);
		if (wText.empty())
			return;

		wstring wOuterHTML;
		unsigned int lastPosition = 0;
		unsigned int numWordsAtStart = sp_vWords->size();

		wstring regPtn = L"[" + WordChars + L"]";
		if (bWithSpaces)
			regPtn.append(L"+");

		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator regit(wText.begin(), wText.end(), wrgx);
		wregex_iterator rend;

		while (regit != rend)
		{
			if (regit->position() != lastPosition)
			{
				wstring skipped = wstring(wText.begin()+lastPosition, wText.begin()+regit->position());
				wOuterHTML.append(skipped);
				if (!bWithSpaces || skipped.find_first_not_of(L' ') != wstring::npos)
				{
					sp_vWords->push_back(wTermDivider);
					AppendTermSpan(wOuterHTML, L"", L"", L"TermDivide", 0);
				}
			}
			lastPosition = regit->position() + regit->length();
			sp_vWords->push_back(chaj::str::wstring_tolower(regit->str()));
			AppendTermSpan(wOuterHTML, regit->str(), sp_vWords->back(), L"Unloaded", 0);
			regit++;
		}
		
		if (sp_vWords->size() == numWordsAtStart)
			return; // no words found, just return

		sp_vWords->push_back(wTermDivider);
		AppendTermSpan(wOuterHTML, L"", L"", L"TermDivide", 0);

		if (lastPosition != wText.size()) // non-word text beyond last found word
			wOuterHTML.append(wText.begin()+lastPosition, wText.end()); // append all remaining

		// transform text node
		SmartCom<IHTMLElement> pNew = CreateElement(pDoc, L"span");
		if (!pNew)
			return;
		SmartCom<IHTMLDOMNode> pNewNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pNew);
		if (!pNewNode)
			return;
		SmartCom<IHTMLDOMNode> pTextNode = GetAlternateInterface<IHTMLDOMTextNode,IHTMLDOMNode>(pText);
		if (!pTextNode)
			return;
		HRESULT hr = pTextNode->replaceNode(pNewNode, pTextNode);
		if (SUCCEEDED(hr))
			HRESULT hr = SetElementOuterHTML(pNew, wOuterHTML);
	}
	void ParseTextNode_AllAtOnce(IHTMLDOMTextNode* pText, IHTMLDocument2* pDoc, unordered_set<wstring>& usetPageTerms)
	{
		if (!pText || !pDoc)
			return;

		wstring wText = GetTextNodeText(pText);
		if (wText.empty())
			return;

		wstring wOuterHTML;
		unsigned int lastPosition = 0;

		SmartCom<IHTMLElement> pNew = CreateElement(pDoc, L"span");
		if (!pNew)
			return;

		SmartCom<IHTMLDOMNode> pNewNode = GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pNew);
		SmartCom<IHTMLDOMNode> pTextNode = GetAlternateInterface<IHTMLDOMTextNode,IHTMLDOMNode>(pText);
		SmartCom<IHTMLElement> pBody = GetBodyFromDoc(pDoc);
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
		unsigned int curPosition = 0;
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

				AppendMultiWordSpan(wOuterHTML, get<0>(current), &get<1>(current), to_wstring(get<2>(current)));
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

		HRESULT hr = pTextNode->replaceNode(pNewNode, pTextNode);
		if (SUCCEEDED(hr))
			hr = chaj::DOM::SetElementOuterHTML(pNew, wOuterHTML);
	}
	void Thread_CachePageMWTerms(LPSTREAM pDocStream, shared_ptr<vector<wstring>> sp_vWords)
	{
		if (!pDocStream || !sp_vWords)
			return;

		const int MAX_MYSQL_INLIST_LEN = 500;

		SmartCom<IHTMLDocument2> pDoc = GetInterfaceFromStream<IHTMLDocument2>(pDocStream);
		if (!pDoc)
			return;

		list<wstring> cWords;			// will hold all words, then have tracked terms removed
		wstring	wInList;				// list of words passed to sql query
		wstring wQuery;					// the sql querystring
		wstring wCanonicalTerm;			// a result's standard form

		for (unsigned int i = 0, j = 0, listLength = 0; i < (*sp_vWords).size(); i += j)
		{
			// this check happens about once for each MAX_MYSQL_INLIST_LEN words on page
			if (bShuttingDown)
				return;

			// create buffer of uncached terms
			wInList.clear();
			for (listLength = 0, j = 0; listLength < MAX_MYSQL_INLIST_LEN && i+j < (*sp_vWords).size(); ++j)
			{
				wstring wRoot = (*sp_vWords)[i+j];
				if (wRoot == wTermDivider)
					continue;

				wstring wTerm = wRoot;
				for (unsigned int k = 1; k <= LWT_MAX_MWTERM_LENGTH-1 && i+j+k < (*sp_vWords).size(); ++k)
				{
					wstring wPiece = (*sp_vWords)[i+j+k];
					if (wPiece == wTermDivider)
						break;

					if (this->bWithSpaces)
						wTerm += L" ";
					wTerm += wPiece;

					if (!wInList.empty())
						wInList.append(L",");

					wInList.append(L"'");
					wInList.append(EscapeSQLQueryValue(wTerm));
					wInList.append(L"'");

					++listLength;
				}
			}

			// cache uncached tracked terms
			if (!wInList.empty())
			{
				wQuery.clear();
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
					wCanonicalTerm = tStmt.field(1).as_string();
					TermRecord rec(tStmt.field(2).as_string());
					rec.wRomanization = tStmt.field(3).as_string();
					rec.wTranslation = tStmt.field(4).as_string();
					cache->insert(make_pair(wCanonicalTerm, rec));
					UpdateCacheMWFragments(wCanonicalTerm);
				}
				tStmt.free_results();
				LeaveCriticalSection(&CS_UseDBConn);
			}
			// at this point all terms in list that are tracked are cached
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
	void Thread_CachePageWords(LPSTREAM pDocStream, shared_ptr<vector<wstring>> sp_vWords)
	{
		if (!pDocStream || !sp_vWords || (*sp_vWords).empty())
			return;

		const int MAX_MYSQL_INLIST_LEN = 500;

		SmartCom<IHTMLDocument2> pDoc = GetInterfaceFromStream<IHTMLDocument2>(pDocStream);
		if (!pDoc)
			return;


		list<wstring> cWords;			// will hold all words, then have tracked terms removed
		wstring	wInList;				// list of words passed to sql query
		wstring wQuery;					// the sql querystring
		wstring wCanonicalTerm;			// a result's standard form

		for (unsigned int i = 0, j = 0, listLength = 0; i < (*sp_vWords).size(); i += j)
		{
			// this check happens about once for each MAX_MYSQL_INLIST_LEN words on page
			if (bShuttingDown)
				return;

			wInList.clear();
			cWords.clear();

			// create list of uncached terms
			for (j = 0, listLength = 0; listLength < MAX_MYSQL_INLIST_LEN && i+j < (*sp_vWords).size(); ++j)
			{
				if ((*sp_vWords)[i+j] == wTermDivider)
					continue;

				cache_it it = cache->find((*sp_vWords)[i+j]);
				if (TermNotCached(it))
				{
					cWords.push_back((*sp_vWords)[i+j]);
					if (!wInList.empty())
						wInList.append(L",");

					wInList.append(L"'");
					wInList.append(EscapeSQLQueryValue(cWords.back()));
					wInList.append(L"'");

					++listLength;
				}
			}

			// cache uncached tracked terms
			if (!wInList.empty())
			{
				wQuery.clear();
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
					wCanonicalTerm = tStmt.field(1).as_string();
					TermRecord rec(tStmt.field(2).as_string());
					rec.wRomanization = tStmt.field(3).as_string();
					rec.wTranslation = tStmt.field(4).as_string();
					cache->insert(make_pair(wCanonicalTerm, rec));
					cWords.remove(wCanonicalTerm);
				}
				tStmt.free_results();
				LeaveCriticalSection(&CS_UseDBConn);
			}
			// at this point all terms in list that are tracked are cached

			for (wstring& word : cWords)
				cache->insert(make_pair(word, TermRecord(L"0")));
		}
	}
	void Thread_GetPageWords(LPSTREAM pDocStream, shared_ptr<vector<wstring>> sp_vWords)
	{
		SmartCom<IHTMLDocument2> pDoc = GetInterfaceFromStream<IHTMLDocument2>(pDocStream);
		if (!pDoc)
			return;

		SmartCom<IDOMTreeWalker> pDocTree = GetDocumentTree(pDoc);
		if (!pDocTree)
			return;

		IDispatch* pNextNodeDisp = nullptr;
		IDispatch* pAfterTextDisp = nullptr;
		IHTMLDOMNode* pNextNode = nullptr;
		IHTMLElement* pElement = nullptr;
		IHTMLDOMTextNode* pText = nullptr;
		bool bAwaitNextElement = false;
		HRESULT hr;

		// loop regex setup
		wstring regPtn = L"[" + WordChars + L"]";
		if (bWithSpaces)
			regPtn.append(L"+");
		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator rend;
		
		// this while loop manages reference counts manually
		long lngType;
		while(SUCCEEDED(pDocTree->nextNode(&pNextNodeDisp)) && pNextNodeDisp)
		{
			pNextNode = GetAlternateInterface<IDispatch,IHTMLDOMNode>(pNextNodeDisp);
			if (pNextNode)
			{
				hr = pNextNode->get_nodeType(&lngType);
				if (SUCCEEDED(hr))
				{
					if (lngType == NODE_TYPE_ELEMENT)
					{
						bAwaitNextElement = false;
						pElement = GetAlternateInterface<IHTMLDOMNode,IHTMLElement>(pNextNode);
						if (pElement)
						{
							wstring wTag = GetTagFromElement(pElement);
							chaj::str::wstring_tolower(wTag);
							if (wTag == L"script" ||
								wTag == L"style" ||
								wTag == L"textarea" ||
								wTag == L"iframe" ||
								wTag == L"noframe")
							{
								bAwaitNextElement = true;
							}
							pElement->Release();
						}
					}
					else if (!bAwaitNextElement && lngType == NODE_TYPE_TEXT)
					{
						dhr(pDocTree->nextNode(&pAfterTextDisp)); // grab the node that follows this text node as a bookmark
						pText = GetAlternateInterface<IHTMLDOMNode,IHTMLDOMTextNode>(pNextNode);
						if (pText)
						{
							MineTextNodeForWords_InsertPlaceHolders(pText, pDoc, sp_vWords);
							pText->Release();
						}
						if (pAfterTextDisp) // kinda hackish, but take whatever node now precedes the bookmark so loop will increment back to bookmark
						{
							pAfterTextDisp->Release();
							hr = pDocTree->previousNode(&pAfterTextDisp);
							if (SUCCEEDED(hr) && pAfterTextDisp)
								pAfterTextDisp->Release();
						}
					}
				}
				pNextNode->Release();
			}
			pNextNodeDisp->Release();
			if (bShuttingDown)
				break;
		}
	}
	IWebBrowser2* GetBrowserFromStream(LPSTREAM pBrowserStream)
	{
		return GetInterfaceFromStream<IWebBrowser2>(pBrowserStream);
	}
	bool StartWordCache_Helper__Thread_OnPageFullyLoaded(thread*& SWordThread, IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords)
	{
		LPSTREAM pDocStream = GetInterfaceStream(IID_IHTMLDocument2, pDoc);
		if (pDocStream)
		{
			SWordThread = new std::thread(&LwtBho::Thread_CachePageWords, this, pDocStream, sp_vWords);
			return true;
		}
		else
			return false;
	}
	bool StartMWCache_Helper__Thread_OnPageFullyLoaded(thread*& MWordThread, IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords)
	{
		LPSTREAM pDocStream = GetInterfaceStream(IID_IHTMLDocument2, pDoc);
		if (pDocStream)
		{
			MWordThread = new std::thread(&LwtBho::Thread_CachePageMWTerms, this, pDocStream, sp_vWords);
			return true;
		}
		else
			return false;
	}
	bool StartPageHighlightThread_Helper__Thread_OnPageFullyLoaded(IHTMLDocument2* pDoc)
	{
		LPSTREAM pDocStream = GetInterfaceStream(IID_IHTMLDocument2, pDoc);
		if (pDocStream)
		{
			cpThreads.push_back(new std::thread(&LwtBho::Thread_HighlightPage, this, pDocStream));
			return true;
		}
		else
			return false;
	}
	void StartPageMining_Helper__Thread_OnPageFullyLoaded(thread*& miner, IHTMLDocument2* pDoc, shared_ptr<vector<wstring>>& sp_vWords)
	{
		LPSTREAM pDocStream = GetInterfaceStream(IID_IHTMLDocument2, pDoc);
		if (pDocStream)
			miner = new std::thread(&LwtBho::Thread_GetPageWords, this, pDocStream, sp_vWords);
	}
	void Thread_OnPageFullyLoaded(LPSTREAM pBrowserStream)
	{
		if (!pBrowserStream)
			return;

		SmartCom<IWebBrowser2> pBrowser = GetBrowserFromStream(pBrowserStream);
		if (!pBrowser)
			return;

		SmartCom<IHTMLDocument2> pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc)
			return;

		AppendCss(pDoc);
		AppendHtmlBlocks(pDoc);
		AppendJavascript(pDoc);

		SmartCom<IHTMLElement> pSetupLink = chaj::DOM::GetElementFromId(L"lwtSetupLink", pDoc);
		if (!pSetupLink) // TODO: add an error log message
			return;
		dhr(pSetupLink->click());

		vector<wstring> vWords;
		shared_ptr<vector<wstring>> sp_vWords = make_shared<vector<wstring>>(vWords);

		thread* miner = nullptr;
		StartPageMining_Helper__Thread_OnPageFullyLoaded(miner, pDoc, sp_vWords);
		if (miner && miner->joinable())
			miner->join();

		// start caching threads
		thread* MWordThread = nullptr;
		thread* SWordThread = nullptr;
		StartMWCache_Helper__Thread_OnPageFullyLoaded(MWordThread, pDoc, sp_vWords);
		StartWordCache_Helper__Thread_OnPageFullyLoaded(SWordThread, pDoc, sp_vWords);

		// wait for caching to complete
		if (MWordThread && MWordThread->joinable())
			MWordThread->join();
		if (SWordThread && SWordThread->joinable())
			SWordThread->join();

		// highlight page with cached terms
		StartPageHighlightThread_Helper__Thread_OnPageFullyLoaded(pDoc);
	}
	void Thread_HighlightPage(LPSTREAM pDocStream)
	{
		if (!pDocStream)
			return;

		unordered_set<wstring> usetPageTerms; // used to track which terms have had an html hidden div record appended

		SmartCom<IHTMLDocument2> pDoc = GetInterfaceFromStream<IHTMLDocument2>(pDocStream);
		if (!pDoc)
			return;

		SmartCom<IHTMLElement> pBody = GetBodyFromDoc(pDoc);
		if (!pBody)
			return;

		SmartCom<IHTMLElementCollection> pCollection = GetAllElementsFromDoc(pDoc);
		if (!pCollection)
			return;

		long lLength = 0;
		HRESULT hr = pCollection->get_length(&lLength);
		if (FAILED(hr) || lLength == 0)
			return;

		SmartCom<IDispatch> pDispElem;
		SmartCom<IHTMLElement> pSpan;
		VARIANT vIndex; VariantInit(&vIndex); vIndex.vt = VT_I4; 
		VARIANT vSubIndex; VariantInit(&vSubIndex); vSubIndex.vt = VT_I4; vSubIndex.lVal = 0;
		VARIANT varNewClassValue; VariantInit(&varNewClassValue); varNewClassValue.vt = VT_BSTR;

		vector<vector<IHTMLElement*>> lwtSpans; // to hold all highlighted words
		vector<IHTMLElement*> lwtRun; // words not separated by non-word parts
		vector<vector<wstring>> lwtTome;
		vector<wstring> lwtPhrase;
		for (long i = 0; i < lLength; ++i)
		{
			vIndex.lVal = i;
			hr = pCollection->item(vIndex, vSubIndex, (IDispatch**)pDispElem);
			if (FAILED(hr) || !pDispElem)
				continue;

			pSpan = GetElementFromDisp(pDispElem);
			if (!pSpan)
				continue;

			if (GetElementClass(pSpan) == L"lwtStatUnloaded")
			{
				lwtPhrase.push_back(GetAttributeValue(pSpan, L"lwtTerm"));
				lwtRun.push_back(pSpan.Relinquish());
			}
			else if (GetElementClass(pSpan) == L"lwtStatTermDivide")
			{
				lwtSpans.push_back(lwtRun);
				lwtRun.clear();
				lwtTome.push_back(lwtPhrase);
				lwtPhrase.clear();
			}
		}

		wstring multiWordTerm;
		stack<TermEntry> found;
		for (unsigned int k = 0; k < lwtTome.size(); ++k)
		{
			for (unsigned int i = 0; i < lwtTome[k].size(); ++i)
			{
				multiWordTerm = lwtTome[k][i];
				for (unsigned int j = 1; i+j < LWT_MAX_MWTERM_LENGTH - 1 && i+j < lwtTome[k].size(); ++j)
				{
					if (bWithSpaces)
						multiWordTerm += L" ";
					multiWordTerm += lwtTome[k][i+j];
					cache_it iter = cache->find(multiWordTerm);
					if (iter != cache->end())
						found.push(make_pair(multiWordTerm, iter->second));
				}
				while (found.size())
				{
					TermEntry termEntry = found.top(); found.pop();
					wstring wOuterHtml;
					AppendMultiWordSpan(wOuterHtml, termEntry.first, &(termEntry.second),
						to_wstring(WordsInTerm(termEntry.first)));
					AppendHTMLBeforeBegin(wOuterHtml, lwtSpans[k][i]);
					if (usetPageTerms.find(termEntry.first) == usetPageTerms.end())
					{
						AppendTermDivRec(pBody, termEntry.first, termEntry.second);
						usetPageTerms.insert(termEntry.first);
					}
				}
				cache_it iter = cache->find(lwtTome[k][i]);
				if (iter != cache->end())
				{
					wstring wOuterHtml;
					AppendSoloWordSpan(wOuterHtml, GetElementInnerText(lwtSpans[k][i]), iter->first, &iter->second);
					HRESULT hr = SetElementOuterHTML(lwtSpans[k][i], wOuterHtml);
					if (usetPageTerms.find(iter->first) == usetPageTerms.end())
					{
						AppendTermDivRec(pBody, iter->first, iter->second);
						usetPageTerms.insert(iter->first);
					}
				}
			}
		}
	}
	static void DeleteHandles()
	{
		for (vector<HANDLE>::size_type i = 0; i < vDelete.size(); ++i)
			DeleteObject(vDelete[i]);
		vDelete.clear();
	}

#pragma warning( push )
#pragma warning( disable : 4100 )
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
#pragma warning ( pop )
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
	void HandleRightClick()
	{
		IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc) return;

		IHTMLEventObj* pEvent = GetEventFromDoc(pDoc);
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
				__noop;//ForcePageParse(pDoc);
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

		SmartCom<IHTMLElement> pCurTermDivRec = GetElementFromId(wCurTermId, pDoc);
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
		SmartCom<IHTMLElement> pCurInfoTerm = GetElementFromId(L"lwtcurinfoterm", pDoc);
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
		SmartCom<IHTMLElement> pLast = GetElementFromId(L"lwtlasthovered", pDoc);
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

		SmartCom<IHTMLElement> pJSCommand = GetElementFromId(L"lwtJSCommand", pDoc);
		if (!pJSCommand) return;

		chaj::DOM::SetAttributeValue(pJSCommand, L"lwtAction", L"");
		chaj::DOM::SetAttributeValue(pJSCommand, L"value", L"");
	}
	bool HandleFillMWTerm(IHTMLDocument2* pDoc)
	{
		if (!pDoc) return false;

		SmartCom<IHTMLElement> pCurInfoTerm = GetElementFromId(L"lwtcurinfoterm", pDoc);
		if (!pCurInfoTerm) return false;

		wstring wCurInfoTerm = GetAttributeValue(pCurInfoTerm, L"lwtterm");
		if (!wCurInfoTerm.size()) return false;

		cache_cit it = cache->cfind(wCurInfoTerm);
		if (TermNotCached(it)) return false;

		SmartCom<IHTMLElement> pTrans = GetElementFromId(L"lwtshowtrans", pDoc);
		SmartCom<IHTMLElement> pRom = GetElementFromId(L"lwtshowtrans", pDoc);
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
	void OnClick()
	{
		SmartCom<IHTMLDocument2> pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc)
			return;

		SmartCom<IHTMLEventObj> pEvent = GetEventFromDoc(pDoc);

		SmartCom<IHTMLElement> pElement = GetClickedElementFromEvent(pEvent);

		LPSTREAM pBrowserStream = GetBrowserStream();
		LPSTREAM pElementStream = GetInterfaceStream(IID_IHTMLElement, pElement);

		if (pBrowserStream && pElementStream)
			std::thread(&LwtBho::Thread_HandleClick, this, pBrowserStream, pElementStream).detach();
	}
	void Thread_HandleClick(LPSTREAM pBrowserStream, LPSTREAM pClickedElementStream)
	{
		if (!pBrowserStream || !pClickedElementStream)
			return;

		SmartCom<IWebBrowser2> pBrowser = GetBrowserFromStream(pBrowserStream);
		SmartCom<IHTMLElement> pElement = GetInterfaceFromStream<IHTMLElement>(pClickedElementStream);

		SmartCom<IHTMLDocument2> pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc) return;

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
	LPSTREAM GetBrowserStream()
	{
		return GetInterfaceStream(IID_IWebBrowser2, pBrowser);
	}
	void OnDocumentComplete(DISPPARAMS* pDispParams)
	{
		if (pDispParams->cArgs >= 2 && pDispParams->rgvarg[1].vt == VT_DISPATCH)
		{
			IDispatch* pCur = pDispParams->rgvarg[1].pdispVal;
			if (pCur == pDispBrowser)
			{
				SmartCom<IHTMLDocument2> pDoc = GetDocumentFromBrowser(pBrowser);
				ReceiveEvents(pDoc);
				LPSTREAM pBrowserStream = GetBrowserStream();
				if (pBrowserStream)
					cpThreads.push_back(new std::thread(&LwtBho::Thread_OnPageFullyLoaded, this, pBrowserStream));
			}
		}
	}
#pragma warning( push )
#pragma warning( disable : 4100 )
	HRESULT _stdcall Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		if (dispidMember == DISPID_HTMLDOCUMENTEVENTS_ONCLICK)
			OnClick();
		else if (dispidMember == DISPID_BEFORENAVIGATE2)
			__noop;
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
#pragma warning( pop )

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
	// Init functions
	void InitializeCriticalSections();

	// Shutdown functions
	void DeleteCriticalSections();
	void WaitForThreadsToTerminate();

	// Database functions
	void ConnectToDatabase();
	void GetLwtActiveTableSetPrefix();

	void LoadCssFile();
	void LoadJavascriptFile();

	/* member variables */

	// static members
	static chaj::DOM::DOMIteratorFilter* s_pFilter;
	static chaj::DOM::DOMIteratorFilter* s_pFullFilter;

	unordered_set<wstring> usetCacheMWFragments;

	// interfaces
	IUnknown* pSite;
	IWebBrowser2* pBrowser;
	IConnectionPoint* pCP;
	IConnectionPoint* pHDEv;
	IDOMTreeWalker* pTW;
	IDispatch* pDispBrowser;

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
	bool bWithSpaces;
	_bstr_t bstrUrl;
	LWTCache* cache;
	unordered_map<wstring, LWTCache*> cacheMap;
	vector<std::thread*> cpThreads;
	int mNumDetachedThreads;
	// wregex
	wregex_iterator rend;
	bool bShuttingDown;
	CRITICAL_SECTION CS_UseDBConn;
	CRITICAL_SECTION CS_General;
};

inline LwtBho::LwtBho()
	:
	bDocumentCompleted(false),
	bShuttingDown(false),
	bWithSpaces(true),
	cache(nullptr),
	Cookie_pHDev(0),
	dwCookie(0),
	hBrowser(nullptr),
	mNumDetachedThreads(0),
	pBrowser(nullptr),
	pCP(nullptr),
	pDispBrowser(nullptr),
	pHDEv(nullptr),
	pSite(nullptr),
	pTW(nullptr),
	ref(0)
{
#ifdef _DEBUG
	mb("Here's your chance to attach debugger...");
#endif

	InitializeCriticalSections();

	// load embedded resources
	LoadJavascriptFile();
	LoadCssFile();

	// initiate database
	ConnectToDatabase();
	GetLwtActiveTableSetPrefix();
	OnTableSetChange(nullptr, false);
}

inline LwtBho::~LwtBho()
{
	bShuttingDown = true;

	WaitForThreadsToTerminate();
	DeleteCriticalSections();

	for (auto& cacheEntry : cacheMap)
		delete cacheEntry.second;
}

#endif