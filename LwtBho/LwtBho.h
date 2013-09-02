#ifndef LwtBho_INCLUDED
#define LwtBho_INCLUDED

#include "LWTCache.h"
#include "C:\\Users\\Luis\\Desktop\\charlie\\prog\\utilities\\resedit\\projects\\resource.h"
#include "assert.h"
#include "dbdriver.h"
#include <Windows.h>
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
#include <vector>
#include <map>
#include <unordered_map>
#include <unordered_set>
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
const INT_PTR CTRL_FORCE_PARSE = static_cast<INT_PTR>(101);
const INT_PTR CTRL_CHANGE_LANG = static_cast<INT_PTR>(102);
const int MWSpanAltroSize = 7;
const wstring wstrNewline(L"&#13;&#10;");

#define nHiddenChunkOpenBookmarkLen 3
#define LWT_MAXWORDLEN 255
wstring wHiddenChunkBookmark(L"~-~");
wstring wPTagBookmark(L"~P~");
wstring wStatIntro(L"lwtStat");
HINSTANCE hInstance;
vector<HANDLE> vDelete;
LONG gref=0;

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
	LwtBho() : conn(), tConn(), tStmt()
	{
#ifdef _DEBUG
		mb("Here's your chance to attach debugger...");
#endif

		pSite = NULL;
		pBrowser = NULL;
		pCP = NULL;
		pHDEv = NULL;
		pNI = NULL;
		pTW = NULL;
		pDispBrowser = NULL;

		hBrowser = NULL;
		dwCookie = NULL;
		dwCookie2 = NULL;
		bDocumentCompleted = false;
		ref = 0;
		LgID = "";
		wLgID = L"";
		tLgID = _T("");
		WordChars = L"";
		SentDelims = L"";
		SentDelimExcepts = L"";
		bWithSpaces = true;
		bstrUrl = "";

		InitializeWregexes();

		pFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTerm);
		pFullFilter = new chaj::DOM::DOMIteratorFilter(&FilterNodes_LWTTermAndSent);

		if (!tConn.connect(L"LWT", L"root", L""))
	    {
			mb(_T("Cannot connect to the Data Source"));
			mb(tConn.last_error());
			return;
		}

		try
		{
			SetCharsetNameOption* scno = new SetCharsetNameOption("utf8");
			conn.set_option(scno);
			conn.connect("learning-with-texts", "127.0.0.1", "root", "", 0);
		}
		catch(...)
		{
			MessageBox(NULL, _T("Lwt Bho unable to connect to MySQL server. Please make sure it is running. Make sure you can access Learning With Texts. If you continue to encounter this message, you can disable this Lwt Bho in your browser add-on options."), _T("60xhfdufio"), MB_OK);
		}

		if (!tStmt.execute_direct(tConn, _T("select StValue from settings where StKey = 'currentlanguage'")))
		{
			mb(_T("Cannot execute query!"));
			mb(tConn.last_error());
			return;
		}
		if (tStmt.fetch_next() == false)
		{
			mb(_T("Now rows returned. Expected db entry."));
			mb(tConn.last_error());
			tStmt.free_results();
			return;
		}
		tLgID = tStmt.field(1).as_string();
		wLgID = chaj::str::to_wstring(tLgID);
		tStmt.free_results();
		OnLanguageChange();
	}
	~LwtBho()
	{
		if (pFilter != NULL)
			delete pFilter;
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
	bool DoExecute(const tstring& errCode)
	{
		bool success = tStmt.execute();
		if (!success)
			DoError(errCode, tStmt.last_error());
		return success;
	}
	bool DoExecuteDirect(const tstring& errCode, const tstring& tSql)
	{
		tStmt.free_results(); // just in case
		bool success = tStmt.execute_direct(tConn, tSql);
		if (!success)
		{
			mb("DoExecuteDirect did not succeed");
			wstring werr = tConn.last_error();
			DoError(errCode, werr);
		}
		return success;
	}
	bool DoPrepare(const tstring& errCode, const tstring& tSql)
	{
		bool success = tStmt.prepare(tConn, tSql);
		if (!success)
			DoError(errCode, tStmt.last_error());
		return success;
	}
	void AddNewTerm(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		wstring wTermEscaped = EscapeSQLQueryValue(wTerm);
		wstring wQuery;
		wQuery.append(L"insert into words (WoLgID, WoText, WoTextLC, WoStatus, WoStatusChanged,  WoTodayScore, WoTomorrowScore, WoRandom) values (");
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
				UpdatePageMWList(wTerm);
				AddNewMWSpans(wTerm, wNewStatus, pDoc);
			}
			else
				UpdatePageSpans(pDoc, wTerm, wNewStatus);
		}
		else
		{
			unordered_map<wstring,TermRecord>::iterator it = cache.find(wTerm);
			if (it != cache.end())
				ChangeTermStatus(wTerm, wNewStatus, pDoc);
			else
				mb(L"Unable to add term. This is a database related issue. Error code: 325nijok", wTerm);
		}
	}
	void AddNewMWSpans(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		HRESULT hr;
		unsigned int uCount = WordsInTerm(wTerm);
		VARIANT_BOOL vBool;
		_bstr_t bFind(wTerm.c_str());
		_bstr_t bstrChar(L"character");
		_bstr_t bstrHowEndToStart(L"EndToStart");
		IHTMLElement* pParent = nullptr;
		long lVal;
		
		wstring out;
		unordered_map<wstring,TermRecord>::const_iterator it = cache.cfind(wTerm);
		assert(it != cache.end());

		IHTMLElement* pBody = chaj::DOM::GetBodyAsElementFromDoc(pDoc);
		assert(pBody);
		AppendTermRecordDiv(pBody, wTerm, it->second);
		pBody->Release();
		AppendMWSpan(out, wTerm, &(it->second), to_wstring(uCount));

		_bstr_t bInsert(out.c_str());

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);

		pRange->findText(bFind, 1000000, 0, &vBool);
		
		while (vBool == VARIANT_TRUE)
		{
			hr = pRange->setEndPoint(bstrHowEndToStart.GetBSTR(), pRange);
			assert(SUCCEEDED(hr));

			IHTMLElement* pSpan = nullptr;
			hr = pRange->parentElement(&pSpan);
			assert(SUCCEEDED(hr));

			while (!chaj::DOM::GetAttributeValue(pSpan, L"lwtterm").size() && pSpan)
			{
				IHTMLElement* pPar = nullptr;
				hr = pSpan->get_parentElement(&pPar);
				assert(SUCCEEDED(hr));

				pSpan->Release();
				pSpan = pPar;
			}

			if (!pSpan)
				return AddNewMWSpans2(wTerm, wNewStatus);

			_bstr_t bWhere(L"beforeBegin");
			hr = pSpan->insertAdjacentHTML(bWhere.GetBSTR(), bInsert.GetBSTR());
			//hr = pRange->pasteHTML(bInsert.GetBSTR());
			assert(SUCCEEDED(hr));

			pRange->move(bstrChar, out.size() + wTerm.size() - 1, &lVal); // todo I'm not sure if this works well, or how it responds to the previous insertion in terms of txtrange cursor position, but I'm hoping, within bounds, that it's reasonable = hack
			pRange->findText(bFind, 1000000, 0, &vBool);
		}
	}
	void AddNewMWSpans2(const wstring& wTerm, const wstring& wNewStatus)
	{
		vector<wstring> vParts;

		if (bWithSpaces)
		{
			int pos = -1; //offset to jive with do loop

			do
			{
				++pos;
				pos = wTerm.find(L" ", pos);
				vParts.push_back(wstring(wTerm, 0, pos));
			} while (pos != wstring::npos);
		}
		else
		{
			for (int i = 0; i < wTerm.size(); ++i)
				vParts.push_back(to_wstring(wTerm.c_str()[i]));
		}


		IDispatch *pRoot = NULL, *pCur = NULL;
		IHTMLElement* pBegin = NULL;
		pTW->get_root(&pRoot);
		if (!pRoot) return;

		pTW->putref_currentNode(pRoot);
		pCur = pRoot;

		do
		{
			IHTMLElement* pElement = GetAlternateInterface<IDispatch, IHTMLElement>(pCur);
			
			if (pElement)
			{
				wstring wCurTerm = GetAttributeValue(pElement, L"lwtTerm");
				if (wCurTerm == vParts[0])
				{
					pBegin = pElement;
					bool bAddHere = false;

					for (int i = 1; i < vParts.size(); ++i)
					{
						pTW->nextNode(&pCur);
						IHTMLElement* pElement2 = GetAlternateInterface<IDispatch,IHTMLElement>(pCur);
						if (!pElement2) break;

						wstring wCurTerm2 = GetAttributeValue(pElement2, L"lwtTerm");
						pElement2->Release();

						if (wCurTerm2 != vParts[i])
							break;

						if (i + 1 == vParts.size())
							bAddHere = true;
					}

					if (bAddHere)
					{
						wstring out;
						out.append(L"<sup><span class=\"lwtStat");
						out += wNewStatus;
						out.append(L"\" lwtTerm=\"");
						out.append(wTerm);
						out.append(L"\">");
						out += to_wstring(vParts.size());
						out.append(L"</span></sup>");
						_bstr_t bOut(out.c_str());
						_bstr_t bWhere(L"beforeBegin");
						pBegin->insertAdjacentHTML(bWhere.GetBSTR(), bOut.GetBSTR());
					}
					pTW->putref_currentNode(pBegin);
				}
				pElement->Release();
			}
			pTW->nextNode(&pCur);
		} while (pCur); 
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

		pRange->findText(bFind, 1000000, 0, &vBool);
		
		while (vBool == VARIANT_TRUE)
		{
			pRange->parentElement(&pParent);
			if (pParent)
			{
				wstring wCurTerm = GetAttributeValue(pParent, L"lwtTerm");
				if (wTerm == wCurTerm)
				{
					HRESULT hr = RemoveNodeFromElement(pParent);
					assert(SUCCEEDED(hr));
#ifdef _DEBUG
					if (FAILED(hr))
						TRACE(L"Unable to properly delete multiword span.");
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
		wQuery.append(L"delete from words where WoTextLC = '");
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
		unordered_map<wstring,TermRecord>::iterator it = cache.find(wTerm);
		assert(it != cache.end());
		(*it).second.wStatus = wNewStatus;
	}
	void InsertOrUpdateCacheItem(const wstring& wTerm, const wstring& wNewStatus)
	{
		unordered_map<wstring,TermRecord>::iterator it = cache.find(wTerm);
		if (it == cache.end())
			cache.insert(unordered_map<wstring,TermRecord>::value_type(wTerm, TermRecord(wNewStatus)));
		else
			(*it).second.wStatus = wNewStatus;
	}
	void ChangeTermStatus(const wstring& wTerm, const wstring& wNewStatus, IHTMLDocument2* pDoc)
	{
		if (wNewStatus == L"0")
			return RemoveTerm(pDoc, wTerm);

		wstring wQuery;
		wQuery.append(L"update words set WoStatus = ");
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
		IHTMLElement* pBody = GetBodyAsElementFromDoc(pDoc);
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
		
			for (int i = 0; i < wstrTerm.size(); ++i)
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
		_bstr_t bWord(wTerm.c_str());
		VARIANT_BOOL vBool;
		pRange->findText(bWord, 1000000, 0, &vBool);
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
	void OnLanguageChange()
	{
		if (LgID == "0")
			return;
		        
		tstring tQuery = _T("select LgRegexpWordCharacters, LgRegexpSplitSentences, COALESCE(LgExceptionsSplitSentences, ''), LgSplitEachChar, COALESCE(LgDict1URI, ''), COALESCE(LgDict2URI, ''), COALESCE(LgGoogleTranslateURI, '') from languages where LgID = ");
		tQuery += tLgID;
		if (!DoExecuteDirect(_T("235ywiehf"), tQuery))
			return;

		if (!tStmt.fetch_next())
			mb("Could not retrieve Word Characters. 249ydhfu.");
		else
		{
			WordChars += tStmt.field(1).as_string();
			SentDelims = tStmt.field(2).as_string();
			SentDelimExcepts = RegexEscape(tStmt.field(3).as_string());
			tStmt.field(4).as_string() == L"0" ? bWithSpaces = true : bWithSpaces = false;
			wstrDict1 = tStmt.field(5).as_string();
			wstrDict2 = tStmt.field(6).as_string();
			wstrGoogleTrans = tStmt.field(7).as_string();
			RegexEscape(SentDelimExcepts);
		}
		tStmt.free_results();
	}
	wstring RegexEscape(const wstring& in)
	{
		return regex_replace(in, wregex(L"([$.?\^*+()])"), L"\\$1");
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
		int i = nOpenTagPos + 1;

		do
		{
			wchar_t wCur = in[i];
			if (wCur == '\"' || wCur == '\'')
			{
				int j = i + 1;
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

		int pos1 = 0, pos2 = 0;
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
	void PopulateChunks(vector<wstring>& chunks, wstring& wFullBody)
	{
		TRACE(L"%s", L"Calling PopulateChunks\n");
				
		if (wFullBody.find(L"iframe") == wstring::npos)
			return SplitChunksByScript(chunks, wFullBody);

		int pos1 = 0, pos2 = 0;
		bool bOpenScriptContext = false; // have seen open <script> but not yet </script>

		wstring wrgxPtn1 = L"< */? *iframe.*?>";
		wregex wrgx1(wrgxPtn1, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
		wregex_iterator regScripts(wFullBody.begin(), wFullBody.end(), wrgx1);

		if (regScripts == rend) // a script-tagless page
			return SplitChunksByScript(chunks, wFullBody);

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
			pos2 = wFullBody.find(wScript, pos1);
			nEnd = FindProperCloseTagPos(wFullBody, pos2);
			if (nEnd-pos2+1 != wScript.size())
				wScript = wstring(wFullBody, pos2, nEnd-pos2+1);

			if (pos1 != pos2) // some content prior to the found script tag
			{
				if (bOpenScriptContext == true)
					chunks.push_back(wstring(wFullBody, pos1, pos2 - pos1));
				else // this is an open script tag we found, what precedes is not javascript
					SplitChunksByScript(chunks, wstring(wFullBody, pos1, pos2 - pos1));
			}
			
			chunks.push_back(wScript);

			pos1 = pos2 + wScript.size();
			++regScripts;
			bOpenScriptContext = !bOpenScriptContext;
			if (bOpenScriptContext && TagIsSelfClosed(wScript))
				bOpenScriptContext = !bOpenScriptContext;
		}
		if (pos1 < wFullBody.size())
			SplitChunksByScript(chunks, wstring(wFullBody, pos1, wFullBody.size()-pos1));

		TRACE(L"%s", L"Leaving PopulateChunks\n");
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
			wcsncmp(chunks[nChunkIndex-1].c_str(), L"<iframe", 7) == 0)
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

	void GetVisibleBodyWithTagBookmarks(wstring& wVisibleBody /*out*/, vector<wstring>& chunks)
	{
		wVisibleBody.clear(); //a single whole coherent result

		bool bIsVisibleChunk; // must be properly set each loop iteration

		for (int i = 0; i < chunks.size(); ++i)
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
					wstring_replaceAll(out, found, L"&#38;");
				else if (wcsicmp(entity.c_str(), L"lt") == 0)
					wstring_replaceAll(out, found, L"&#60;");
				else
					wstring_replaceAll(out, found, to_wstring(entities[entity]));
			}
			else
			{
				if (entity.c_str()[0] == '#')
				{
					wstring wNum(entity, 1, entity.size() - 1);
					UINT uNum = stoul(wNum);
					wstring wstrChar(to_wstring((wchar_t)uNum));
					wstring_replaceAll(out, found, wstrChar);
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

	void UpdatePageMWList(const wstring& wMWTerm)
	{
		if (bWithSpaces)
		{
			int pos = -1; //offset to jive with do loop

			do
			{
				++pos;
				pos = wMWTerm.find(L" ", pos);
				usetPageMWList.insert(wstring(wMWTerm, 0, pos));
			} while (pos != wstring::npos);
		}
		else
		{
			for (int i = 1; i <= wMWTerm.size(); ++i)
				usetPageMWList.insert(wstring(wMWTerm, 0, i));
		}
	}

	// GetUncacedSubset expects candidate words to be in lowercase already
	void GetUncachedSubset(vector<wstring>& vListUncached, const vector<wstring>& vCandidateList)
	{
		for (wstr_c_it itWord = vCandidateList.begin(), itWordEnd = vCandidateList.end(); itWord != itWordEnd; ++itWord)
		{
			unordered_map<wstring,TermRecord>::const_iterator it = cache.find(*itWord);
			if (it == cache.end())
			{
				//mb("not found");
				vListUncached.push_back(*itWord);
			}
			//else
			//{
			//	mb("the item");
			//	mb(*itWord);
			//	mb(to_wstring((*itWord).size()));
			//	UpdatePageMWList(*itWord);
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
		for (int i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			for (int j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				wstring vTerm;
				int nSkipCount = 0;

				for (int k = j; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
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

					//assert(cache.count(tokens[i][j]) == 1);
					//TermRecord rec = cache[tokens[i][j]];
					//if (rec.nTermsBeyond == 0)
					//	break;
				}
			}
		}
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

	//				assert(cache.count(wWord) == 1);
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

		for (int i = 0; i < vWordsNoDups.size(); ++i)
		{
			unordered_map<wstring,TermRecord>::iterator it = cache.find(vWordsNoDups[i]);
			if (it == cache.end())
				continue;
			
			wstring wQuery = L"select exists(select 1 from words where WoTextLC like '%";
			wQuery += vWordsNoDups[i];
			if (bWithSpaces == true)
				wQuery.append(L" %'");
			else
				wQuery.append(L"%'");
			wQuery.append(L" LIMIT 1)");
			//mb(wQuery);
			if (!DoExecuteDirect(_T("677idjfjf"), wQuery))
				break;

			if (tStmt.fetch_next())
				it->second.nTermsBeyond = tStmt.field(1).as_long() > 0 ? 1 : 0;

			tStmt.free_results();
		}
	}
	void UpdateCaches(const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling UpdateCaches\n");
		EnsureRecordEntryForEachWord(tknCanonical);
		//UpdateTerminationCache(vWords);
		EnsureRecordEntryForEachMWTerm(tknCanonical);
		TRACE(L"%s", L"Leaving UpdateCaches\n");
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
		wWordSansBookmarksWrgxPtn.append(L"(?:");
		wWordSansBookmarksWrgxPtn.append(wHiddenChunkBookmark);
		wWordSansBookmarksWrgxPtn.append(L"|");
		wWordSansBookmarksWrgxPtn.append(wPTagBookmark);
		wWordSansBookmarksWrgxPtn.append(L")[0-9]+(?:");
		wWordSansBookmarksWrgxPtn.append(wHiddenChunkBookmark);
		wWordSansBookmarksWrgxPtn.append(L"|");
		wWordSansBookmarksWrgxPtn.append(wPTagBookmark);
		wWordSansBookmarksWrgxPtn.append(L")");
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
		for (int i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			for (int j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				words.push_back(WordSansBookmarks(tokens[i][j]));
			}
		}
	}
	//void SentenceWordsToWordsMinusBookmarks(vector<wstring>& words, const vector<vector<wstring>>& vSentencesVWords)
	//{
	//	for (int i = 0; i < vSentencesVWords.size(); ++i)
	//	{
	//		for (int j = 0; j < vSentencesVWords[i].size(); ++j)
	//			words.push_back(WordSansBookmarks(vSentencesVWords[i][j]));
	//	}
	//}
	wstring EscapeSQLQueryValue(const wstring& wQuery)
	{
		TRACE(L"Entering EscapeSQLQueryValue: const overload\n");
		
		wstring out(wQuery);
		EscapeSQLQueryValue(out);
		return out;

		TRACE(L"Leaving EscapeSQLQueryValue: const overload\n");
	}
	wstring& EscapeSQLQueryValue(wstring& wQuery)
	{
		TRACE(L"Entering EscapeSQLQueryValue\n");

		int pos = wQuery.find('\''), pos2;
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

		return wQuery;
		TRACE(L"Leaving EscapeSQLQueryValue\n");
	}
	void InsertCacheItems(int nAtATime, list<wstring>& lstItems, bool bIsMultiWord = false)
	{
		TRACE(L"%s", L"Calling InsertCacheItems\n");
		int MAX_FIELD_WIDTH = 255;
		bool bWithList = false;
		vector<wstring> vQueries;
		list<wstring> lstCopy(lstItems);

		list<wstring>::iterator lend = lstCopy.end();
		list<wstring>::iterator iter = lstCopy.begin();
		for (int i = 0; i < lstCopy.size(); i += nAtATime)
		{
			TRACE(L"Creating query #%i\n", i+1);
			wstring wInList;

			for (int j = i; j < lstCopy.size() && j < i + nAtATime; ++j)
			{
				wInList.append(L"'");
				wInList.append(EscapeSQLQueryValue(*iter));
				wInList.append(L"',");
				++iter;
			}

			wInList.append(L"'~~z~x~~Q_a_dummy_value'");
			wstring wQuery;
			wQuery.reserve(nAtATime * MAX_FIELD_WIDTH);
			wQuery.append(L"select WoTextLC, WoStatus, COALESCE(WoRomanization, ''), COALESCE(WoTranslation, '') from words where WoLgID = ");
			wQuery.append(wLgID);
			wQuery.append(L" AND WoTextLC in (");
			wQuery.append(wInList);
			wQuery.append(L")");
			vQueries.push_back(wQuery);
		}

		TRACE(L"%i queries created\n", vQueries.size());
		for (int i = 0; i < vQueries.size(); ++i)
		{
			DoExecuteDirect(_T("1612isjdlfij"), vQueries[i]);
			TRACE(L"Executed query #%i\n", i + 1);
			while (tStmt.fetch_next())
			{
				wstring wLC = tStmt.field(1).as_string();
				TermRecord rec(tStmt.field(2).as_string());
				rec.wRomanization = tStmt.field(3).as_string();
				rec.wTranslation = tStmt.field(4).as_string();
				cache.insert(unordered_map<wstring,TermRecord>::value_type(wLC, rec));
				if (bIsMultiWord == true)
					UpdatePageMWList(wLC);
				else
					lstItems.remove(wLC);
			}
			tStmt.free_results();
		}
		TRACE(L"%s", L"Leaving InsertCacheItems\n");
	}
	void GetUncachedTokens(list<wstring>& out, const TokenStruct& tokens)
	{
		TRACE(L"%s", L"Calling GetUncachedTokens\n");
		for (int i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested)
			{
				for (int j = 0; j < tokens[i].size(); ++j)
				{
					if (tokens[i].isWordDigested(j))
					{
						unordered_map<wstring,TermRecord>::const_iterator it = cache.find(tokens[i][j]);
						if (it == cache.end())
							out.push_back(tokens[i][j]);
					}
				}
			}
		}
		TRACE(L"%s", L"Leaving GetUncachedTokens\n");
	}
	void EnsureRecordEntryForEachWord(const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling EnsureRecordEntryForEachWord\n");
		list<wstring> lstUncachedWords;
		GetUncachedTokens(lstUncachedWords, tknCanonical);

		int MAX_MYSQL_INLIST_LEN = 2000;

		InsertCacheItems(MAX_MYSQL_INLIST_LEN, lstUncachedWords);

		for(list<wstring>::iterator itList = lstUncachedWords.begin(), end = lstUncachedWords.end(); itList != end; ++itList)
			cache.insert(unordered_map<wstring,TermRecord>::value_type(*itList, TermRecord(L"0")));
		TRACE(L"%s", L"Leaving EnsureRecordEntryForEachWord\n");
	}
	void test5(vector<wstring>& v)
	{
		for (int i = 0; i < v.size(); ++i)
			mb(v[i]);
	}
	void EnsureRecordEntryForEachMWTerm(const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling EnsureRecordEntryForEachMWTerm\n");
		list<wstring> lstTerms;
		ParseSentencesToMWTerms(lstTerms, tknCanonical);

		//vector<wstring> vUncachedMWTerms;
		//GetUncachedSubset(vUncachedMWTerms, vTerms);

		int nMaxMWTermsPerInList = 500;
		InsertCacheItems(nMaxMWTermsPerInList, lstTerms, true);
		TRACE(L"%s", L"Leaving EnsureRecordEntryForEachMWTerm\n");
	}
	//void EnsureRecordEntryForEachMWTerm(const vector<vector<wstring>>& vSentencesVWords)
	//{
	//	vector<wstring> vTerms;
	//	ParseSentencesToMWTermsOnlyIf_SansBookmarks(vTerms, vSentencesVWords);

	//	vector<wstring> vUncachedMWTerms;
	//	GetUncachedSubset(vUncachedMWTerms, vTerms);

	//	int nMaxMWTermsPerInList = 500;
	//	InsertCacheItems(nMaxMWTermsPerInList, vUncachedMWTerms);
	//}

	/*
		GetBodyContent
		Input conditions: pBody is NULL, pDoc is not NULL
		Exit condition: pDoc pointer equal to entrance value and has not had Release called
		If successful, pBody non-NULL
	*/
	void GetBodyContent(wstring& wBody, IHTMLElement*& pBody, IHTMLDocument2* pDoc)
	{
		TRACE(L"%s", L"Calling GetBodyContent\n");
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
				pHDEv->Unadvise(dwCookie2);
				pHDEv->Release();
			}
		
			pHDEv = pHDEv2;
			hr = pHDEv->Advise(reinterpret_cast<IDispatch*>(this), &dwCookie2);
		}
		else
			pHDEv2->Release();
		if (FAILED(hr))
			mb("Error. Will not perceive mouse clicks on this page. Will try to annotate anyway.", "1000duwksm");

		hr = pDoc->get_body(&pBody);
		if (FAILED(hr))
		{
			mb(L"get_body gave an error", L"132asdfedds");
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
		
		wBody = to_wstring(bstrBodyContent.GetBSTR());
		TRACE(L"%s", L"Leaving GetBodyContent\n");
	}
	void PSTS_WordLevel_spaceless(TokenStruct& tokens, const wstring& in)
	{
		int pos1 = 0, pos2 = 0;
		Token_Sentence ts(true);

		// pattern: (?:(?:~-~#~-~)*[aWordChar])(?:~-~#~-~|[aWordChar])*
		// regex pattern regPtn ensures that it returns a true word (contains at least one lwt defined word character)
		// regPtns also ensures that if ~ is considered a valid wordChar, we won't end up at - and choke on a bookmark tag,
		// i.e. it eats whole bookmarks first at each comparison point, and it why the second appearance of [aWordChar]
		// is not [aWordChar]+
		wstring regPtn = L"[";
		regPtn += WordChars;
		regPtn.append(L"]");

		wregex wrgx(regPtn, regex_constants::ECMAScript);
		wregex_iterator regit(in.begin(), in.end(), wrgx);
		wregex_iterator rend;
		int i = tokens.size();

		while (regit != rend)
		{
			wstring wWord = regit->str();
			assert(pos1 != wstring::npos && pos1 >= 0 && pos1 < in.size());
			pos2 = in.find(wWord, pos1);
			assert(pos2 != wstring::npos); // regex pattern just found should be found again
			if (pos2 != pos1) // non-word data found in sentence
				ts.push_back(wstring(in, pos1, pos2-pos1), false);
			pos1 = pos2 + wWord.size();
			
			ts.push_back(wWord);

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
	void PSTS_WordLevel(TokenStruct& tokens, const wstring& in)
	{
		int pos1 = 0, pos2 = 0;
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
			regPtn.append(L"])(?:");
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
				for (int i = 0; i < wWord.size(); ++i)
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
		int pos1 = 0, pos2 = 0;

		// pattern: (?:acceptable_sub_sequence|[^a_sentence_separactor_char])+
		//wstring rgxPtn = L"(?:";
		//rgxPtn += SentDelimExcepts + L"|[^";
		//rgxPtn += SentDelims + L"])+";
		wstring rgxPtn = L"[^";
		rgxPtn += SentDelims + L"]+";
		if (tokens.size() == 28)
			int idh = 3;
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

			PSTS_WordLevel(tokens, wSentence);

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
		int pos1 = 0, pos2 = 0;
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
			catch (const std::invalid_argument& ia)
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
			catch (const std::invalid_argument& ia)
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
		out.append(L"</sup>");
	}
	void AppendWithSegues(wstring& out, const wstring& wAsAppended, const wstring& wTermCanonical, const TermRecord* pRec)
	{
		AppendTermIntro(out, wTermCanonical, pRec);
		out.append(wAsAppended);
		AppendTermAltro(out);
	}
	void AppendTermIntro2(wstring& out, const wstring& wTermCanonical, TermRecord* pRec)
	{
		out.append(L"<span class=\"lwtStat");
		out += pRec->wStatus;
		out.append(L"\" lwtTerm=\"");
		out.append(wTermCanonical);
		out.append(L"\"");
		
		out.append(L" title=\"");
		out.append(pRec->wRomanization);
		out.append(wstrNewline);
		out.append(pRec->wTranslation);
		out.append(L"\"");

		out.append(L">");
	}
	void AppendTermIntro(wstring& out, const wstring& wTermCanonical, const TermRecord* pRec)
	{
		out.append(L"<span class=\"lwtStat");
		out += pRec->wStatus;
		out.append(L"\" lwtTerm=\"");
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
	void AppendInfoBoxJavascript(IHTMLDocument2* pDoc, IHTMLElement* pBody)
	{
		IHTMLElement* pHead = chaj::DOM::GetHeadAsElementFromDoc(pDoc);
		IHTMLElement* pScript = nullptr;
		BSTR bstrScript = SysAllocString(L"script");
		pDoc->createElement(bstrScript, &pScript);
		SysFreeString(bstrScript);

		wstring out;

		//out.append(
			//L"window.lwtdict1 = getElementById('lwtdict1').getAttribute('src');"
			//L"window.lwtdict2 = getElementById('lwtdict2').getAttribute('src');"
			//L"window.lwtgoogletrans = getElementById('lwtgoogletrans').getAttribute('src');"
			//);
		out.append(L"function lwtcheckinner() {alert('could call inserted javascript!');}");
		out.append(
L"function totalLeftOffset(element)"
L"{"
L"	var amount = 0;"
L"	while (element != null)"
L"	{"
L"		amount += element.offsetLeft;"
L"		element = element.offsetParent;"
L"	}"
L"	return amount;"
L"}"
L""
L"function totalTopOffset(element)"
L"{"
L"	var amount = 0;"
L"	while (element != null)"
L"	{"
L"		amount += element.offsetTop;"
L"		element = element.offsetParent;"
L"	}"
L"	return amount;"
L"}"
L""
L"function fixPageXY(e)"
L"{"
L"	if (e.pageX == null && e.clientX != null )"
L"	{"
L"		var html = document.documentElement;"
L"		var body = document.body;"
L"		e.pageX = e.clientX + (html.scrollLeft || body && body.scrollLeft || 0);"
L"		e.pageX -= html.clientLeft || 0;"
L"		e.pageY = e.clientY + (html.scrollTop || body && body.scrollTop || 0);"
L"		e.pageY -= html.clientTop || 0;"
L"	}"
L"}"
L""
L"function setSelection(elem, lastterm, curTerm)"
L"{"
L"	elem.id = 'lwtcursel';"
L"	lastterm.setAttribute('lwtcursel', curTerm);"
L"	elem.className = elem.className + ' lwtSel';"
L"}"
L"function getSelection()"
L"{"
L"	var elem = document.getElementById('lwtcursel');"
L"	if (elem === null)"
L"		return '';"
L"	return elem.getAttribute('lwtcursel');"
L"}"
L""
L"function removeSelection()"
L"{"
L"	var elem = document.getElementById('lwtcursel');"
L"	if (elem == null)"
L"	{"
L"		return;"
L"	}"
L"	elem.id = '';"
L"	document.getElementById('lwtlasthovered').setAttribute('lwtcursel','');"
L"	var curClass = elem.className;"
L"	var lastSpace = curClass.lastIndexOf(' ');"
L"	if (lastSpace >= 0)"
L"	{"
L"		var putClass = curClass.substr(0, lastSpace + 1);"
L"		elem.className = putClass;"
L"	}"
L"}"
L""
L"function XOnElement(x, elem)"
L"{"
L"	if (x < totalLeftOffset(elem) || x >= totalLeftOffset(elem) + elem.offsetWidth)"
L"	{"
L"		return false;"
L"	}"
L"	return true;"
L"}"
L""
L"function YOnElement(y, elem)"
L"{"
L"	if (y < totalTopOffset(elem) || y >= totalTopOffset(elem) + elem.offsetHeight)"
L"	{"
L"		return false;"
L"	}"
L"	return true;"
L"}"
L""
L"function XYOnElement(x, y, elem)"
L"{"
L"	return (XOnElement(x, elem) && YOnElement(y, elem));"
L"}"
L""
L"function lwtcontextexit()"
L"{"
L"	removeSelection();"
L"	CloseOpenDialogs();"
L"}"
L""
L"function lwtdivmout(e)"
L"{"
L"	fixPageXY(e);"
L"	var statbox = window.curStatbox;"
L"	var bVal = XYOnElement(e.pageX, e.pageY, statbox);"
L"	if (bVal == false)"
L"	{"
L"		lwtmout(e);"
L"	}"
L"}"
L""
L"function lwtmout(e)"
L"{"
L"	fixPageXY(e);"
L"	var statbox = window.curStatbox;"
L"	if (!XYOnElement(e.pageX, e.pageY, statbox))"
L"	{"
L"		lwtcontextexit();"
L"	}"
L"}"
L""
L"function CloseOpenDialogs()"
L"{"
L"	window.curStatbox.style.display = 'none';"
L"}"
L""
L"function lwtkeypress()"
L"{"
L"	if (document.getElementById('lwtlasthovered').getAttribute('lwtcursel') == '')"
L"	{"
L"		return;"
L"	}"
L"	var e = window.event;"
L"	if (e.ctrlKey || e.altKey || e.shiftKey || e.metaKey)"
L"	{"
L"		return;"
L"		alert('in lwtkeypress return!');"
L"	}"
L"	switch(e.keyCode)"
L"	{"
L"		case 49:"
L"			document.getElementById('lwtsetstat1').click();"
L"			break;"
L"		case 50:"
L"			document.getElementById('lwtsetstat2').click();"
L"			break;"
L"		case 51:"
L"			document.getElementById('lwtsetstat3').click();"
L"			break;"
L"		case 52:"
L"			document.getElementById('lwtsetstat4').click();"
L"			break;"
L"		case 53:"
L"			document.getElementById('lwtsetstat5').click();"
L"			break;"
L"		case 87:"
L"			document.getElementById('lwtsetstat99').click();"
L"			break;"
L"		case 73:"
L"			document.getElementById('lwtsetstat98').click();"
L"			break;"
L"		case 85:"
L"			document.getElementById('lwtsetstat0').click();"
L"			break;"
L"		case 27:"
L"			CloseOpenDialogs();"
L"			break;"
L"	}"
L"}"
L""
L"function ExtrapLink(elem, linkForm, curTerm)"
L"{"
L"	if (elem == null) {return;}"
L"	var extrap = 'javascript:void(0);';"
L"	if (linkForm.indexOf('*') == 0)"
L"	{"
L"		extrap = linkForm.substr(1, linkForm.length);"
L"		extrap = extrap.replace('###', curTerm);"
L"		elem.setAttribute('target', '_blank');"
L"	}"
L"	else if (linkForm.indexOf('###') >= 0)"
L"	{"
L"		extrap = linkForm.replace('###', curTerm);"
L"		elem.setAttribute('target', '_blank');"
L"	}"
L"	elem.setAttribute('href', extrap);"
L"}"
L""
L"function lwtmover(whichid, e, origin)"
L"{"
L"	removeSelection();"
L"	var divRec = document.getElementById(whichid);"
L"	if (divRec == null) {alert('could not locate divRec');}"
L"	var curTerm = divRec.getAttribute('lwtterm');"
L"	var lastterm = document.getElementById('lwtlasthovered');"
L"	if (lastterm == null) {alert('could not loccated lasthovered field');}"
L"	lastterm.setAttribute('lwtterm', curTerm);"
L"	lastterm.setAttribute('lwtstat', divRec.getAttribute('lwtstat'));"
L"	setSelection(origin, lastterm, curTerm);"
L"	fixPageXY(e);"
L"	lwtshowinlinestat(e, curTerm, origin);"
L"	var infobox = document.getElementById('lwtinfobox');"
L"	if (infobox == null) {alert('could not locate infobox to render popup');}"
L"	infobox.style.left = (e.pageX + 10) + 'px';"
L"	infobox.style.top = e.pageY + 'px';"
L"	document.getElementById('lwtinfoboxterm').innerText = curTerm;"
L"	document.getElementById('lwtinfoboxtrans').innerText = divRec.getAttribute('lwttrans');"
L"	document.getElementById('lwtinfoboxrom').innerText = divRec.getAttribute('lwtrom');"
L"}"
L""
L"function lwtshowinlinestat(e, curTerm, origin)"
L"{"
L"	var statbox = null;"
L"	var curStat = '';"
L""
L"	if (window.mwTermBegin != null)"
L"	{"
L"		statbox = document.getElementById('lwtInlineMWEndPopup');"
L"		if (window.keydownEventAttached == true)"
L"		{"
L"			document.detachEvent('onkeydown', lwtkeypress);"
L"			window.keydownEventAttached = false;"
L"		}"
L"	}"
L"	else"
L"	{"
L"		statbox = document.getElementById('lwtinlinestat');"
L"		curStat = document.getElementById('lwtcursel').className;"
L"		ExtrapLink(document.getElementById('lwtextrapdict1'), document.getElementById('lwtdict1').getAttribute('src'), curTerm);"
L"		ExtrapLink(document.getElementById('lwtextrapdict2'), document.getElementById('lwtdict2').getAttribute('src'), curTerm);"
L"		ExtrapLink(document.getElementById('lwtextrapgoogletrans'), document.getElementById('lwtgoogletrans').getAttribute('src'), curTerm);"
L"		if (window.keydownEventAttached != true)"
L"		{"
L"			document.attachEvent('onkeydown', lwtkeypress);"
L"			window.keydownEventAttached = true;"
L"		}"
L"	}"
L""
L"	if (statbox == null)"
L"	{"
L"		alert('could not locate inline status change popup');"
L"		return;"
L"	}"
L""
L"	window.curStatbox = statbox;"
L""
L"	var posElem = origin;"
L"	var altElem = e.srcElement;"
L"	var inlineTop2 = (totalTopOffset(posElem) + posElem.offsetHeight - 2);"
L"	var inlineTop = totalTopOffset(posElem);"
L"	if (curStat.indexOf('9') >= 0)"
L"	{"
L"		inlineTop -= 2;"
L"	}"
L""
L""
L"	statbox.style.left = totalLeftOffset(posElem) + 'px';"
L"	document.getElementById('lwtTermTrans').style.width = posElem.offsetWidth + 'px';"
L"	document.getElementById('lwtTermTrans2').style.width = posElem.offsetWidth + 'px';"
L"	statbox.style.top = inlineTop + 'px';"
L""
L"	var inlineBottom = inlineTop + statbox.offsetHeight;"
L"	statbox.style.display = 'block';"
L"	if (inlineBottom > window.pageYOffset + window.innerHeight)"
L"		statbox.scrollIntoView(false);"
L"}"
L""
L"function traverseDomTree_NextNodeByTagName(elem, aTagName)"
L"{"
L"	if (elem.hasChildNodes() == true)"
L"	{"
L"		if (elem.firstChild.tagName == aTagName)"
L"			return elem.firstChild;"
L"		else"
L"			return traverseDomTree_NextNodeByTagName(elem.firstChild, aTagName);"
L"	}"
L""
L"	var possNextElem = elem.nextSibling;"
L"	if (possNextElem == null)"
L"	{"
L"		while (elem.parentNode.nextSibling == null)"
L"		{"
L"			elem = elem.parentNode;"
L"			if (elem == null)"
L"				return null;"
L"		}"
L""
L"		if (elem.parentNode.nextSibling.tagName == aTagName)"
L"			return elem.parentNode.nextSibling;"
L"		else"
L"			return traverseDomTree_NextNodeByTagName(elem.parentNode.nextSibling, aTagName);"
L"	}"
L""
L"	if (possNextElem.tagName == aTagName)"
L"		return possNextElem;"
L"	else"
L"		return traverseDomTree_NextNodeByTagName(possNextElem, aTagName);"
L"}"
L""
L"function beginMW()"
L"{"
L"	window.mwTermBegin = document.getElementById('lwtcursel');"
L"	lwtcontextexit();"
L"}"
L""
L"function cancelMW()"
L"{"
L"	window.mwTermBegin = null;"
L"	lwtcontextexit();"
L"}"
L""
L"function captureMW(newStatus)"
L"{"
L"	var newTerm = getChosenMWTerm(document.getElementById('lwtcursel'));"
L"	cancelMW();"
L"	if (newTerm != '')"
L"	{"
L"		var lastterm = document.getElementById('lwtlasthovered');"
L"		if (lastterm == null) {alert('could not locate lasthovered field');}"
L"		lastterm.setAttribute('lwtterm', newTerm);"
L"		lastterm.setAttribute('lwtstat', '0');"
L""
L"		switch(newStatus)"
L"		{"
L"			case 49:"
L"				document.getElementById('lwtsetstat1').click();"
L"				break;"
L"			case 50:"
L"				document.getElementById('lwtsetstat2').click();"
L"				break;"
L"			case 51:"
L"				document.getElementById('lwtsetstat3').click();"
L"				break;"
L"			case 52:"
L"				document.getElementById('lwtsetstat4').click();"
L"				break;"
L"			case 53:"
L"				document.getElementById('lwtsetstat5').click();"
L"				break;"
L"			case 87:"
L"				document.getElementById('lwtsetstat99').click();"
L"				break;"
L"			case 73:"
L"				document.getElementById('lwtsetstat98').click();"
L"				break;"
L"		}"
L"	}"
L"	else"
L"		alert('Not a valid composite term selection.');"
L"}"
L""
L"function getChosenMWTerm(elemLastPart)"
L"{"
L"	var bWithSpaces = false;"
L"	if (document.getElementById('lwtwithspaces').getAttribute('value') == 'yes')"
L"		bWithSpaces = true;"
L""
L"	var curTerm = window.mwTermBegin.getAttribute('lwtterm');"
L"	if (curTerm.length <= 0)"
L"		return '';"
L"	var chosenMWTerm = curTerm;"
L""
L"	var elem = window.mwTermBegin;"
L"	for (var i = 0; i < 8; i++)"
L"	{"
L"		elem = traverseDomTree_NextNodeByTagName(elem, 'SPAN');"
L""
L"		if (elem == null)"
L"			return '';"
L""
L"		curTerm = elem.getAttribute('lwtterm');"
L"		if (curTerm.length <= 0)"
L"			i--;"
L"		else"
L"		{"
L"			if (bWithSpaces == true)"
L"				chosenMWTerm += ' ';"
L"			"
L"			chosenMWTerm += curTerm;"
L"			alert(chosenMWTerm);"
L""
L"			if (elem == elemLastPart)"
L"				return chosenMWTerm;"
L"		}"
L"	}"
L""
L"	return '';"
L"}"
L""
L"function getPossibleMWTermParts(elem)"
L"{"
L"	var mwList = new Array(8);"
L"	for (var i = 0; i < 8; i++)"
L"	{"
L"		mwList[i] = '';"
L"	}"
L""
L"	var curTerm = elem.getAttribute('lwtterm');"
L"		"
L"	if (curTerm.length <= 0)"
L"		return mwList;"
L""
L"	for (0; i < 8; i++)"
L"	{"
L"		elem = traverseDomTree_NextNodeByTagName(elem, 'SPAN');"
L""
L"		if (elem == null)"
L"			return mwList;"
L""
L"		if (elem.className == 'lwtsent')"
L"			return mwList;"
L""
L"		curTerm = elem.getAttribute('lwtterm');"
L"		if (curTerm.length <= 0)"
L"			i--;"
L"		else"
L"			mwList[i] = curTerm;"
L"	}"
L""
L"	return mwList;"
L"}"
L"function getMWTermPartsAccrued(elem)"
L"{"
L"	var parts = getMWTermParts(elem);"
L"	var partsAccrued = new Array(8);"
L""
L"	for (var i = 0; i < 8; i++)"
L"	{"
L"		mwList[i] = '';"
L"	}"
L""
L"	var bWithSpaces = false;"
L"	if (document.getElementById('lwtwithspaces').getAttribute('value') == 'yes')"
L"		bWithSpaces = true;"
L""
L"	var curTerm = elem.getAttribute('lwtterm');"
L"		"
L"	if (curTerm.length <= 0)"
L"		return mwList;"
L""
L"	var accruedTerm = curTerm;"
L""
L"	for (i = 0; i < 8 && parts[i] != ''; i++)"
L"	{"
L"		if (bWithSpaces == true)"
L"			accruedTerm += ' ';"
L""
L"		accruedTerm += parts[i];"
L"		partsAccrued[i] = accruedTerm;"
L"	}"
L""
L"	return partsAccrued;"
L"	"
L"}"

			);


		chaj::DOM::SetAttributeValue(pScript, L"type", L"text/javascript");
		chaj::DOM::SetElementInnerText(pScript, out);

		bool DoHead = false;
		if (pHead && DoHead)
		{
			IHTMLDOMNode* pHeadNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pHead);
			IHTMLDOMNode* pScriptNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pScript);
			IHTMLDOMNode* pNewScriptNode;
			HRESULT hr = pHeadNode->appendChild(pScriptNode, &pNewScriptNode);
			if (FAILED(hr))
				TRACE(L"Unable to append javacript child. 23328ijf");
			pNewScriptNode->Release();
			pScriptNode->Release();
			pHeadNode->Release();

			pHead->Release();
		}
		if (pBody && !DoHead)
		{
			IHTMLDOMNode* pBodyNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pBody);
			IHTMLDOMNode* pScriptNode = chaj::COM::GetAlternateInterface<IHTMLElement,IHTMLDOMNode>(pScript);
			IHTMLDOMNode* pNewScriptNode;
			HRESULT hr = pBodyNode->appendChild(pScriptNode, &pNewScriptNode);
			pNewScriptNode->Release();
			pScriptNode->Release();
			pBodyNode->Release();

			pHead->Release();
		}
		else
			TRACE(L"Unable to append lwt javascript properly to head");
	}
	/*
	AppendAnnotatedContent_MaintainBookmarks
	Input Condition: all words in vSentencesVWords are in the cache container
	*/
	void AppendAnnotatedContent_MaintainBookmarks
	(wstring& out, const TokenStruct& tokens, const TokenStruct& tknWithout, const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling AppendAnnotatedContent_MaintainBookmarks\n");
		for (int i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
			{
				out.append(tokens[i][0]);
				continue;
			}

			out.append(L"<span class=\"lwtSent\"></span>");

			for (int j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
				{
					out.append(tokens[i][j]);
					continue;
				}

				unordered_map<wstring,TermRecord>::const_iterator it = cache.find(tknCanonical[i][j]);
				assert(it != cache.end()); // there should always be a result
			
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
					if (usetPageMWList.count(curTerm) != 0) //this terms starts some MWTerm
					{

						stack<mwVals> stkMWTerms;
						int nSkipCount = 0;
						// gather a stack of term statuses for tracked multiword terms that start at this word
						for (int k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
						{
							if (tokens[i].isWordDigested(k) == false)
							{
								++nSkipCount;
								continue;
							}

							if (bWithSpaces == true)
								curTerm.append(L" ");
							curTerm += tknCanonical[i][k];

							if (usetPageMWList.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
								break;

							unordered_map<wstring,TermRecord>::const_iterator it = cache.find(curTerm);
							if (it != cache.end())
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
	//			unordered_map<wstring,TermRecord>::iterator it = cache.find(curWord);
	//			assert(it != cache.end()); // there should always be a result
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

	//				if (usetPageMWList.count(curWord) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
	//					break;

	//				unordered_map<wstring,TermRecord>::iterator it = cache.find(curWord);
	//				if (it != cache.end())
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
	//			unordered_map<wstring,TermRecord>::iterator it = cache.find(curWord);
	//			assert(it != cache.end()); // there should always be a result
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

	//				if (usetPageMWList.count(curWord) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
	//					break;

	//				unordered_map<wstring,TermRecord>::iterator it = cache.find(curWord);
	//				if (it != cache.end())
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
				++pos;
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
		TRACE(L"%s", L"Calling GetTokensWithoutBookmarks\n");
		for (int i = 0; i < tokens.size(); ++i)
		{
			Token_Sentence ts(tokens[i].bDigested);

			for (int j = 0; j < tokens[i].size(); ++j)
			{
				ts.push_back(WordSansBookmarks(tokens[i][j]), tokens[i].isWordDigested(j));
			}

			tokensWithout.push_back(ts);
		}
		TRACE(L"%s", L"Leaving GetTokensWithoutBookmarks\n");
	}
	void GetTokensCanonical(TokenStruct& tokensCanonical, const TokenStruct& tokensWithout)
	{
		TRACE(L"%s", L"Calling GetTokensCanonical\n");
		for (int i = 0; i < tokensWithout.size(); ++i)
		{
			Token_Sentence ts(tokensWithout[i].bDigested);

			for (int j = 0; j < tokensWithout[i].size(); ++j)
			{
				ts.push_back(chaj::str::wstring_tolower(tokensWithout[i][j]), tokensWithout[i].isWordDigested(j));
			}

			tokensCanonical.push_back(ts);
		}
		TRACE(L"%s", L"Leaving GetTokensCanonical\n");
	}
	/*	
		!!! External dependencies: Adding New MWSpans relies on the filtering of class global TreeWalker pTW
		SetDocTreeWalker
	*/
	void SetDocTreeWalker(IHTMLDocument2* pDoc, IHTMLElement* pRoot = NULL)
	{
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
		IHTMLElement* pBody = GetBodyAsElementFromDoc(pDoc);
		HRESULT hr = pBody->put_innerHTML(bNewBodyContent.GetBSTR());
		pBody->Release();
	}
	void ExpandChunkEntities(vector<wstring>& chunks)
	{
		for (int i = 0; i < chunks.size(); ++i)
			ReplaceHTMLEntitiesWithChars(chunks[i]);
	}
	void NewParse(IHTMLDocument2* pDoc)
	{

		IHTMLTxtRange* pRange = GetBodyTxtRangeFromDoc(pDoc);
		if (!pRange)
			return;

		BSTR bstrHTMLText;

		wstring wstrText = HTMLTxtRange_get_text(pRange);
		pRange->get_htmlText(&bstrHTMLText);
		wstring wstrHTMLText(bstrHTMLText);
		SysFreeString(bstrHTMLText);
		int i = 0;

		IHTMLElement* pBody = NULL;
		wstring wsBodyContent = L"";
		GetBodyContent(wsBodyContent, pBody, pDoc);	
		if (pBody == NULL)
			return;

		// separate html into chunks, each chunk fully one of intra-tag content or extra-tag content:  |<p>|This is |<u>|such|</u>| a good book!|</p>|
		chunks.clear();
		PopulateChunks(chunks, wsBodyContent);
		//SplitChunksByScript(chunks, wsBodyContent);

		// condense non visible data down to placeholders
		wstring wVisibleBody;
		GetVisibleBodyWithTagBookmarks(wVisibleBody, chunks);

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
		ReconstructDoc(tokens, tknSansBookmarks, tknCanonical, pRange, wstrText, wstrHTMLText, pDoc, pBody);	

#ifdef _DEBUG
		IHTMLElement* pBody2 = NULL;
		wstring wsBodyContent2 = L"";
		GetBodyContent(wsBodyContent2, pBody2, pDoc);
		pBody2->Release();
#endif

		pBody->Release();
	}
	void AppendTermRecordDiv(IHTMLElement* pBody, const wstring& wstrTerm, const TermRecord& rec)
	{
		wstring out;
		out.append(L"<div id=\"lwt");
		out.append(to_wstring(rec.uIdent));
		out.append(L"\" lwtstat=\"");
		out.append(rec.wStatus);
		out.append(L"\" lwtterm=\"");
		out.append(wstrTerm);
		out.append(L"\" lwttrans=\"");
		out.append(rec.wTranslation);
		out.append(L"\" lwtrom=\"");
		out.append(rec.wRomanization);
		out.append(L"\" />");
		chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
	}
	void AppendTermRecordDivs(IHTMLElement* pBody)
	{
		unordered_set<wstring> termSet;
		GetSetOfTerms(termSet, tknCanonical);

		for (auto it = termSet.cbegin(); it != termSet.cend(); ++it)
		{
			AppendTermRecordDiv(pBody, *it, cache.find(*it)->second);
		}
	}
	void AppendPopupTemplate(IHTMLElement* pBody)
	{
		wstring out;
		out.append(L"<div id=\"lwtinfobox\" style=\"display:none;position:absolute;background-color:white;color:black;\">"
			L"Term: <span id=\"lwtinfoboxterm\"></span><br />"
			L"Romanization: <span id=\"lwtinfoboxrom\"></span><br />"
			L"Translation: <span id=\"lwtinfoboxtrans\"></span>"
			L"</div>");
		chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
	}
	void AppendPopup_HtmlChangeStatus(IHTMLElement* pBody)
	{
		return AppendPopup_HtmlChangeStatus2(pBody);
		wstring out;
			out.append(
			L"<div id=\"lwtinlinestat\" style=\"display:none;position:absolute;\" onmouseout=\"lwtdivmout(event);\">"
			L"	<span class=\"lwtStat1\" id=\"lwtsetstat1\" lwtstat=\"1\" style=\"border:solid black;left-padding:5px;display:inline-block;width:160px\" lwtstatchange=\"true\"> 1 </span><br />"
			L"	<span class=\"lwtStat2\" id=\"lwtsetstat2\" lwtstat=\"2\" style=\"border:solid black;left-padding:5px;display:inline-block;width:140px\" lwtstatchange=\"true\"> 2 </span><br />"
			L"	<span class=\"lwtStat3\" id=\"lwtsetstat3\" lwtstat=\"3\" style=\"border:solid black;left-padding:5px;display:inline-block;width:120px\" lwtstatchange=\"true\"> 3 </span><br />"
			L"	<span class=\"lwtStat4\" id=\"lwtsetstat4\" lwtstat=\"4\" style=\"border:solid black;left-padding:5px;display:inline-block;width:100px\" lwtstatchange=\"true\"> 4 </span><br />"
			L"	<span class=\"lwtStat5\" id=\"lwtsetstat5\" lwtstat=\"5\" style=\"border:solid black;left-padding:5px;display:inline-block;width:80px\" lwtstatchange=\"true\"> 5 </span><br />"
			L"	<span class=\"lwtStat99\" id=\"lwtsetstat99\" lwtstat=\"99\" style=\"border:solid black;left-padding:5px;display:inline-block;width:60px\" lwtstatchange=\"true\" title=\"Well Known\"> Well </span><br />"
			L"	<span class=\"lwtStat98\" id=\"lwtsetstat98\" lwtstat=\"98\" style=\"border:solid black;left-padding:5px;display:inline-block;width:40px\" lwtstatchange=\"true\" title=\"Ignore\"> Ign </span><br />"
			L"	<span class=\"lwtStat0\" id=\"lwtsetstat0\" lwtstat=\"0\" style=\"border:solid black;left-padding:5px;display:inline-block;width:20px\" lwtstatchange=\"true\" title=\"Untrack\"> U </span><br />"
			L"</div>"
			L"<div id=\"lwtlasthovered\" style=\"display:none;\" lwtterm=\"\"></div>"
			);
		chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
	}
	void AppendPopup_HtmlChangeStatus2(IHTMLElement* pBody)
	{
		wstring out;
			out.append(
			L"<div id=\"lwtinlinestat\" style=\"display:none;position:absolute;\" onmouseout=\"lwtdivmout(event);\">"
			L"<span id=\"lwtTermTrans\">&nbsp;</span><br />"
			L"<span class=\"lwtStat1\" id=\"lwtsetstat1\" lwtstat=\"1\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center\" lwtstatchange=\"true\">1</span>"
			L"<span class=\"lwtStat2\" id=\"lwtsetstat2\" lwtstat=\"2\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">2</span>"
			L"<span class=\"lwtStat3\" id=\"lwtsetstat3\" lwtstat=\"3\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">3</span>"
			L"<span class=\"lwtStat4\" id=\"lwtsetstat4\" lwtstat=\"4\" style=\"border-top:solid black 1px;border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:15px;text-align:center;\" lwtstatchange=\"true\">4</span>"
			L"<span class=\"lwtStat5\" id=\"lwtsetstat5\" lwtstat=\"5\" style=\"border:solid black 1px;display:inline-block;width:15px\;text-align:center;\" lwtstatchange=\"true\">5</span><br />"
			L"<span class=\"lwtStat99\" id=\"lwtsetstat99\" lwtstat=\"99\" style=\"border-left:solid black 1px;border-bottom: solid 1px #CCFFCC;display:inline-block;width:26px;text-align:center;\" lwtstatchange=\"true\" title=\"Well Known\">WK</span>"
			L"<span class=\"lwtStat98\" id=\"lwtsetstat98\" lwtstat=\"98\" style=\"border-left:solid black 1px;display:inline-block;width:26px;text-align:center;\" lwtstatchange=\"true\" title=\"Ignore\">Ig</span>"
			L"<span class=\"lwtStat0\" id=\"lwtsetstat0\" lwtstat=\"0\" style=\"border-left:solid black 1px;border-right:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:25px;text-align:center;\" lwtstatchange=\"true\" title=\"Untrack\">Un</span><br />"
			L"<span id=\"lwtmwstart\" style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:79;text-align:center;background-color:white;\" onclick=\"beginMW();\" mwbuffer=\"true\">Begin MW</span><br />"
			L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 1\"><a id=\"lwtextrapdict1\" href=\"\">D1</a></span>"
			L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;display:inline-block;width:26px;text-align:center;background-color:white;\" title=\"Dictionary 2\"><a id=\"lwtextrapdict2\" href=\"\">D2</a></span>"
			L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:25px;text-align:center;background-color:white;\" title=\"Google Translate Term\"><a id=\"lwtextrapgoogletrans\" href=\"\">GT</a></span>"
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
			L"<span class=\"lwtStat5\" lwtstat=\"5\" style=\"border:solid black 1px;display:inline-block;width:15px\;text-align:center;\" onclick=\"captureMW(53);\">5</span><br />"
			L"<span class=\"lwtStat99\" lwtstat=\"99\" style=\"border-left:solid black 1px;border-bottom: solid 1px #CCFFCC;display:inline-block;width:39px;text-align:center;\" onclick=\"captureMW(87);\" title=\"Well Known\">Well</span>"
			L"<span class=\"lwtStat98\" lwtstat=\"98\" style=\"border-left:solid black 1px;border-right:solid black 1px;display:inline-block;width:39px;text-align:center;\" onclick=\"captureMW(73);\" title=\"Ignore\">Ignore</span><br />"
			L"<span style=\"border-left:solid black 1px;border-bottom:solid black 1px;border-right:solid black 1px;display:inline-block;width:78;text-align:center;background-color:white;\" onclick=\"cancelMW();\" mwbuffer=\"true\">Cancel MW</span>"
			L"</div>"
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

		chaj::DOM::AppendHTMLBeforeEnd(out, pBody);
	}
	void GetSetOfTerms(unordered_set<wstring>& out, const TokenStruct& tknCanonical)
	{
		TRACE(L"%s", L"Calling GetSetOfTerms\n");
		for (int i = 0; i < tknCanonical.size(); ++i)
		{
			if (tknCanonical[i].bDigested)
			{
				for (int j = 0; j < tknCanonical[i].size(); ++j)
				{
					if (tknCanonical[i].isWordDigested(j))
					{
						unordered_map<wstring,TermRecord>::const_iterator it = cache.find(tknCanonical[i][j]);
						assert(it != cache.end());
						out.insert(tknCanonical[i][j]);

						if (it->second.nTermsBeyond != 0) //the term might participate in multiword terms
						{
							wstring curTerm = tknCanonical[i][j];
							if (usetPageMWList.count(curTerm) != 0) //this terms starts some MWTerm
							{
								int nSkipCount = 0;
								// gather a stack of term statuses for tracked multiword terms that start at this word
								for (int k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
								{
									if (tokens[i].isWordDigested(k) == false)
									{
										++nSkipCount;
										continue;
									}

									if (bWithSpaces == true)
										curTerm.append(L" ");
									curTerm += tknCanonical[i][k];

									if (usetPageMWList.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
										break;

									unordered_map<wstring,TermRecord>::const_iterator it = cache.find(curTerm);
									if (it != cache.end())
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
	void ReconstructDoc(const TokenStruct& tokens, const TokenStruct& tknWithout, const TokenStruct& tknCanonical, IHTMLTxtRange* pRange, wstring wstrText, wstring wstrHTMLText, IHTMLDocument2* pDoc, IHTMLElement* pBody)
	{
		HRESULT hr;
		VARIANT_BOOL vBool;
		wstring out;
		int lwtid = 1;

		wstring intro(L".lwtStat0 {background-color: #ADDFFF} .lwtStat1 {background-color: #F5B8A9} .lwtStat2 {background-color: #F5CCA9} .lwtStat3 {background-color: #F5E1A9} .lwtStat4 {background-color: #F5F3A9} .lwtStat5 {background-color: #DDFFDD} .lwtStat99 {background-color: #F8F8F8;border-bottom: solid 2px #CCFFCC} .lwtStat98 {background-color: #F8F8F8;border-bottom: dashed 1px #000000} #lwtcursel {border-left: solid black 1px;border-right: solid black 1px;border-top:solid black 1px}");
		hr = AppendStylesToDoc(intro, pDoc);
		if (FAILED(hr))
		{
			TRACE(L"Leaving ReconstructDoc with error: could not append styles.\n");
			return;
		}

		BSTR bstrHowEndToStart = SysAllocString(L"EndToStart");
		BSTR bstrSentenceMarker = SysAllocString(L"<span class=\"lwtSent\" lwtsent=\"lwtsent\"></span>");
		hr = pRange->setEndPoint(bstrHowEndToStart, pRange);

		for (int i = 0; i < tokens.size(); ++i)
		{
			if (tokens[i].bDigested == false)
				continue;

			hr = pRange->pasteHTML(bstrSentenceMarker);
			assert(hr == S_OK);
			
			for (int j = 0; j < tokens[i].size(); ++j)
			{
				if (tokens[i].isWordDigested(j) == false)
					continue;

				BSTR bWord = SysAllocString(tknWithout[i][j].c_str());
				pRange->findText(bWord, 1000000, 0, &vBool);
				SysFreeString(bWord);
				if (vBool == VARIANT_FALSE)
				{
					TRACE(L"Warning: Could not locate word to replace it: %s\n", tknWithout[i][j].c_str());
					continue;
				}

				unordered_map<wstring,TermRecord>::const_iterator it = cache.find(tknCanonical[i][j]);
				assert(it != cache.end()); // there should always be a result

				if (it->second.nTermsBeyond != 0) //the term might participate in multiword terms
				{
					wstring curTerm = tknCanonical[i][j];
					if (usetPageMWList.count(curTerm) != 0) //this terms starts some MWTerm
					{
						stack<mwVals> stkMWTerms;
						int nSkipCount = 0;
						// gather a stack of term statuses for tracked multiword terms that start at this word
						for (int k = j + 1; k < j + 9 + nSkipCount && k < tokens[i].size(); ++k)
						{
							if (tokens[i].isWordDigested(k) == false)
							{
								++nSkipCount;
								continue;
							}

							if (bWithSpaces == true)
								curTerm.append(L" ");
							curTerm += tknCanonical[i][k];

							if (usetPageMWList.count(curTerm) == 0) //this, and any further MWterms with this base are not in this page, even if they exist in db and/or cache
								break;

							unordered_map<wstring,TermRecord>::const_iterator it = cache.cfind(curTerm);
							if (it != cache.end())
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
					AppendSoloWordSpan(out, tknWithout[i][j], tknCanonical[i][j], &(it->second));
				else
				{
					wstring out(tokens[i][j]);
					RestoreAllBookmarks(out, chunks);
					TRACE(L"%ls ~~~ %ls ~~~ %ls\n", out.c_str(), tokens[i][j].c_str(), tknWithout[i][j].c_str());
				}
				BSTR bNew = SysAllocString(out.c_str());
				out.clear();
				hr = pRange->pasteHTML(bNew);
				assert(hr == S_OK);
				SysFreeString(bNew);
			}
		}

		hr = pRange->pasteHTML(bstrSentenceMarker);
		assert(hr == S_OK);

		AppendTermRecordDivs(pBody);
		AppendPopupTemplate(pBody);
		AppendPopup_HtmlChangeStatus(pBody);
		/* specifically last after all elements placed */ AppendInfoBoxJavascript(pDoc, pBody);

		SysFreeString(bstrHowEndToStart);
		SysFreeString(bstrSentenceMarker);
	}
	void ParsePage(IHTMLDocument2* pDoc)
	{
		return NewParse(pDoc);
		TRACE(L"%s", L"Calling ParsePage\n");
		if (pDoc == NULL)
			return;

		IHTMLElement* pBody = NULL;
		wstring wsBodyContent = L"";
		GetBodyContent(wsBodyContent, pBody, pDoc);	
		if (pBody == NULL)
			return;

		// separate html into chunks, each chunk fully one of intra-tag content or extra-tag content:  |<p>|This is |<u>|such|</u>| a good book!|</p>|
		chunks.clear();
		PopulateChunks(chunks, wsBodyContent);
		//SplitChunksByScript(chunks, wsBodyContent);

		// condense non visible data down to placeholders
		wstring wVisibleBody;
		GetVisibleBodyWithTagBookmarks(wVisibleBody, chunks);

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

		wstring theNewsBodyContent(L"<style>span.lwtStat0 {background-color: #ADDFFF} span.lwtStat1 {background-color: #F5B8A9} span.lwtStat2 {background-color: #F5CCA9} span.lwtStat3 {background-color: #F5E1A9} span.lwtStat4 {background-color: #F5F3A9} span.lwtStat5 {background-color: #DDFFDD} span.lwtStat99 {background-color: #F8F8F8;border-bottom: solid 2px #CCFFCC} span.lwtStat98 {background-color: #F8F8F8;border-bottom: dashed 1px #000000}</style>");
		
		AppendAnnotatedContent_MaintainBookmarks(theNewsBodyContent, tokens, tknSansBookmarks, tknCanonical);

		ReplaceNecessaryCharsWithHTMLEntities(theNewsBodyContent);

		RestoreAllBookmarks(theNewsBodyContent, chunks);

		// replace body text with added annotations
		BSTR bNewsBodyContent = SysAllocStringLen(theNewsBodyContent.data(), theNewsBodyContent.size());
		pBody->put_innerHTML(bNewsBodyContent);
		SysFreeString(bNewsBodyContent);
		pBody->Release();
	TRACE(L"%s", L"Leaving ParsePage\n");
	}
	bool TokensReassembledProperly(const TokenStruct& ts, const wstring& orig)
	{
		wstring result;

		for(int i = 0; i < ts.size(); ++i)
		{
			for(int j = 0; j < ts[i].size(); ++j)
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
			if (pSite != NULL)
				pSite->Release();
			if (pBrowser != NULL)
				pBrowser->Release();
			if (pDispBrowser != NULL)
				pDispBrowser->Release();
			if (pCP != NULL)
			{
				if (dwCookie != NULL)
				{
					hr = pCP->Unadvise(dwCookie);
					if (FAILED(hr))
						mb("failure upon attempt to unadvise", "1484soideif");
				}
				pCP->Release();
			}
			if (pTW != NULL)
				pTW->Release();
			return S_OK;
		}

		pUnkSite->AddRef();
		if (pSite != NULL)
			pSite->Release();
		if (pBrowser != NULL)
			pBrowser->Release();
		if (pDispBrowser != NULL)
			pDispBrowser->Release();
		if (pCP != NULL)
		{
			if (dwCookie != NULL)
				pCP->Unadvise(dwCookie);
			pCP->Release();
		}
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

	void OnDocumentComplete(IHTMLDocument2* pDoc)
	{
		TRACE(L"%s", L"Calling OnDocumentComplete\n");
		SetDocTreeWalker(pDoc);
		ParsePage(pDoc);
		TRACE(L"%s", L"Leaving OnDocumentComplete\n");
	}
	static void DeleteHandles()
	{
		for (int i = 0; i < vDelete.size(); ++i)
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
	static INT_PTR CALLBACK DlgProc_ChangeStatus(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		switch(uMsg)
		{
			case WM_INITDIALOG:
			{
				HBITMAP hBmp;
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS1));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON1),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS2));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON2),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS3));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON3),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS4));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON4),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS5));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON5),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS0));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON6),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS99));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON7),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				hBmp = LoadBitmap(hInstance, MAKEINTRESOURCE(IDB_STATUS98));
				vDelete.push_back(hBmp);
				SendMessage(GetDlgItem(hwndDlg, IDC_BUTTON8),BM_SETIMAGE,(WPARAM)IMAGE_BITMAP,(LPARAM)hBmp);
				
				if (lParam)
				{
					MainDlgStruct* pMDS = reinterpret_cast<MainDlgStruct*>(lParam);
					vector<wstring>& vTerms = pMDS->Terms;
					if (vTerms.size() > 0)
					{
						for (int i = 0; i < vTerms.size(); ++i)
						{
							SendMessageW(GetDlgItem(hwndDlg, IDC_COMBO1), CB_ADDSTRING, NULL, reinterpret_cast<LPARAM>(vTerms[i].c_str()));
						}

						SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_SETCURSEL, (WPARAM)0, (LPARAM)0);
					}
				}
				return (INT_PTR) TRUE;
			}
			case WM_COMMAND:
			{
				switch(HIWORD(wParam))
				{
					case BN_CLICKED:
					{
						DlgResult* pDR;

						int wp = LOWORD(wParam);
						if (wp == IDC_STATIC_LINK)
						{
							DeleteHandles();
							pDR = new DlgResult;
							pDR->nStatus = 100;
							EndDialog(hwndDlg, (INT_PTR)pDR);
							break;
						}
						if (wp == IDC_BUTTON1 || wp == IDC_BUTTON2 || wp == IDC_BUTTON3 || wp == IDC_BUTTON4 || wp == IDC_BUTTON5 || wp == IDC_BUTTON6 || wp == IDC_BUTTON7 || wp == IDC_BUTTON8)
						{
							wchar_t wszEntry[LWT_MAXWORDLEN+1];
							int nCurSel = (int)SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_GETCURSEL, 0, 0);
							SendMessage(GetDlgItem(hwndDlg, IDC_COMBO1), CB_GETLBTEXT, (WPARAM)nCurSel, reinterpret_cast<LPARAM>(wszEntry));
							pDR = new DlgResult;
							pDR->wstrTerm = wszEntry;		
						}

						switch(LOWORD(wParam))
						{
							case IDC_BUTTON1:
							{
								
								DeleteHandles();
								pDR->nStatus = 1;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON2:
							{
								DeleteHandles();
								pDR->nStatus = 2;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON3:
							{
								DeleteHandles();
								pDR->nStatus = 3;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON4:
							{
								DeleteHandles();
								pDR->nStatus = 4;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON5:
							{
								DeleteHandles();
								pDR->nStatus = 5;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON6:
							{
								DeleteHandles();
								pDR->nStatus = 0;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON7:
							{
								DeleteHandles();
								pDR->nStatus = 99;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
							case IDC_BUTTON8:
							{
								DeleteHandles();
								pDR->nStatus = 98;
								EndDialog(hwndDlg, (INT_PTR)pDR);
								break;
							}
						}
					}
					case CBN_SELCHANGE:
					{
						wchar_t wszEntry[LWT_MAXWORDLEN+1];
						int nCurSel = (int)SendMessage(reinterpret_cast<HWND>(lParam), CB_GETCURSEL, 0, 0);
						SendMessage(reinterpret_cast<HWND>(lParam), CB_GETLBTEXT, (WPARAM)nCurSel, reinterpret_cast<LPARAM>(wszEntry));
						SetWindowText(GetDlgItem(hwndDlg, IDC_EDIT1), wszEntry);
						break;
					}
				}
				return 0;
			}
			case WM_CLOSE:
			{
				DeleteHandles();
				DlgResult* pDR = new DlgResult;
				pDR->nStatus = 100;
				EndDialog(hwndDlg, (INT_PTR)pDR);
			}
		}
		return (INT_PTR) FALSE;
	}
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

		IHTMLElement* pBody = GetBodyAsElementFromDoc(pDoc);
		DOMIteratorFilter filter(&FilterNodes_LWTTerm);
		pNI = GetNodeIteratorWithFilter(pDoc, pBody, dynamic_cast<IDispatch*>(&filter));
		pBody->Release();
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
		TRACE(L"Calling SetEventReturnFalse\n");
		VARIANT varRetVal; VariantInit(&varRetVal); varRetVal.vt = VT_BOOL; varRetVal.boolVal = VARIANT_FALSE;
		HRESULT hr = pEvent->put_returnValue(varRetVal);
		VariantClear(&varRetVal);
		if (FAILED(hr))
			TRACE(L"Encountered unexpected hresult in SetEventReturnFalse.\n");
		TRACE(L"Leaving SetEventReturnFalse\n");
		return hr;
	}
	inline bool TermInCache(const wstring& wTerm)
	{
		return cache.find(wTerm) != cache.end();
		//unordered_map<wstring,TermRecord>::iterator it = cache.find(wTerm);
		//return it != cache.end();
	}
	void HandleClick()
	{
		IHTMLDocument2* pDoc = GetDocumentFromBrowser(pBrowser);
		if (!pDoc) return;

		IHTMLEventObj* pEvent = GetEventFromDocument(pDoc);
		if (!pEvent)
		{
			pDoc->Release();
			return;
		}

		IHTMLElement* pElement = GetClickedElementFromEvent(pEvent, pDoc);
		if (!pElement)
		{
			pEvent->Release();
			pDoc->Release();
			return;
		}

		wstring wStatChange = GetAttributeValue(pElement, L"lwtstatchange");
		if (wStatChange.size())
		{
			pEvent->Release();
			IHTMLElement* pLast = GetElementFromId(L"lwtlasthovered", pDoc);
			wstring wLast = GetAttributeValue(pLast, L"lwtterm");
			assert(wLast != L"");

			wstring wCurStat = GetAttributeValue(pLast, L"lwtstat");
			wstring wNewStat = GetAttributeValue(pElement, L"lwtstat");
			if (wCurStat == L"0")
				AddNewTerm(wLast, wNewStat, pDoc);
			else
				ChangeTermStatus(wLast, wNewStat, pDoc);
			
			SetAttributeValue(pLast, L"lwtstat", wNewStat);

			pLast->Release();
			pElement->Release();
			pDoc->Release();
		}
		else if (GetAttributeValue(pElement, L"value").size())
		{
		}
		else
		{

			wstring wTerm = GetAttributeValue(pElement, L"lwtTerm");
			if (wTerm.size() <= 0)
				return;

			//_bstr_t bstrStat;
			BSTR bstrStat;
			HRESULT hr = pElement->get_className(&bstrStat);//bstrStat.GetAddress());
			if (bstrStat == NULL)
				return;

			wstring wOldStatus(bstrStat);
			SysFreeString(bstrStat);
			wstring wOldStatNum(wstring(wOldStatus, wStatIntro.size()));
			bool isAddOp = false;
			if (wOldStatNum == L"0")
				isAddOp = true;


			MainDlgStruct mds(ElementPartOfLink(pElement));
			DlgResult* pDR = NULL;
			int res;
			if (IsMultiwordTerm(wTerm))
			{
				if (mds.bOnLink)
					pDR = (DlgResult*)DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG2), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus));
				else
					pDR = (DlgResult*)DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG1), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus));
				if ((int)pDR == -1)
				{
					mb("Unable to render dialog box, or other related error.", "2232idhf");
					return;
				}

				res = pDR->nStatus;
				if (res != 100)
				{
					assert(res == 0 || res == 1 || res == 2 || res == 3 || res == 4 || res == 5 || res == 98 || res == 99);
					ChangeTermStatus(wTerm, to_wstring(res), pDoc);
					if (mds.bOnLink)
						SetEventReturnFalse(pEvent);
				}
			}
			else
			{
				GetDropdownTermList(mds.Terms, pElement);
				if (mds.bOnLink)
					pDR = (DlgResult*)DialogBoxParam(hInstance, MAKEINTRESOURCE(IDD_DIALOG12), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus), (LPARAM)&mds);
				else
					pDR = (DlgResult*)DialogBoxParam(hInstance, MAKEINTRESOURCE(IDD_DIALOG11), hBrowser, reinterpret_cast<DLGPROC>(DlgProc_ChangeStatus), (LPARAM)&mds);

				if ((int)pDR == -1)
				{
					mb("Unable to render dialog box, or other related error.", "2232idhf");
					return;
				}

				res = pDR->nStatus;
				if (res != 100)
				{
					assert(res == 0 || res == 1 || res == 2 || res == 3 || res == 4 || res == 5 || res == 98 || res == 99);
					if (IsMultiwordTerm(pDR->wstrTerm))
					{
						if (!TermInCache(pDR->wstrTerm))
							AddNewTerm(pDR->wstrTerm, to_wstring(res), pDoc);
						else
							mb("Term is already defined. Isn't it?");
					}
					else
					{
						if (isAddOp == true)
							AddNewTerm(wTerm, to_wstring(res), pDoc);
						else
							ChangeTermStatus(wTerm, to_wstring(res), pDoc);
					}

					VARIANT varRetVal; VariantInit(&varRetVal); varRetVal.vt = VT_BOOL; varRetVal.boolVal = VARIANT_FALSE;
					pEvent->put_returnValue(varRetVal);
					VariantClear(&varRetVal);
				}
			}

			pElement->Release();
			pEvent->Release();
			pDoc->Release();
			delete pDR;
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
	IHTMLDocument2* GetDocumentFromBrowser(IWebBrowser2* pBrowser)
	{
		TRACE(L"%s", L"Entering GetDocumentFromBrowser\n");

		IDispatch* pDisp = nullptr;
		HRESULT hr = pBrowser->get_Document(&pDisp);
		if(FAILED(hr))
		{
			mb(L"Get Document gave an error", L"2804wishgio");
			TRACE(L"%s", L"Leaving GetDocumentFromBrowser with error result\n");
			return nullptr;
		}

		IHTMLDocument2* pDoc = nullptr;
		hr = pDisp->QueryInterface(IID_IHTMLDocument2, reinterpret_cast<void**>(&pDoc));
		pDisp->Release();
		if (FAILED(hr))
		{
			MessageBox(NULL, L"Retrieve IHTMLDoc2 gave an error", L"2813suhfgdju", MB_OK);
			TRACE(L"%s", L"Leaving GetDocumentFromBrowser with error result\n");
			return nullptr;
		}

		TRACE(L"%s", L"Leaving GetDocumentFromBrowser\n");
		return pDoc;
	}
	IHTMLEventObj* GetEventFromDocument(IHTMLDocument2* pDoc)
	{
		TRACE(L"%s", L"Calling GetEventFromDocument\n");

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

		TRACE(L"%s", L"Leaving GetEventFromDocument\n");
		return event;
	}
	IHTMLElement* GetClickedElementFromEvent(IHTMLEventObj* pEvent, IHTMLDocument2* pDoc)
	{
		TRACE(L"%s", L"Entering GetClickedElementFromEvent\n");

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

		TRACE(L"%s", L"Leaving GetClickedElementFromEvent\n");
		return pElement;
	}
	HRESULT _stdcall Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
	{
		if (dispidMember == DISPID_HTMLDOCUMENTEVENTS_ONCLICK)
		{
			TRACE(L"Handling DISPID_HTMLDOCUMENTEVENTS_ONCLICK");
			HandleClick();
			TRACE(L"Handled DISPID_HTMLDOCUMENTEVENTS_ONCLICK");
		}
		if (dispidMember == DISPID_HTMLDOCUMENTEVENTS_ONCONTEXTMENU)
		{
			TRACE(L"Handling DISPID_HTMLDOCUMENTEVENTS_ONCONTEXTMENU");
			HandleRightClick();
			TRACE(L"Handled DISPID_HTMLDOCUMENTEVENTS_ONCONTEXTMENU");
		}
		//if (dispidMember == DISPID_DOWNLOADBEGIN)
		//if (dispidMember == DISPID_DOWNLOADCOMPLETE)
		//if (dispidMember == DISPID_NAVIGATECOMPLETE2)
		//if (dispidMember == DISPID_ONQUIT)
		//if (dispidMember == DISPID_BEFORENAVIGATE2)
		if (dispidMember == DISPID_DOCUMENTCOMPLETE)
		{
			//mb("document complete");
			IDispatch* pCur = pDispParams->rgvarg[1].pdispVal;
			if (pCur == pDispBrowser)
			{
				IDispatch* pDisp = NULL;
				HRESULT hr = pBrowser->get_Document(&pDisp);
				if(FAILED(hr))
				{
					mb(L"Get Document gave an error", L"378wishgio");
					return E_UNEXPECTED;
				}

				IHTMLDocument2* pDoc = NULL;
				hr = pDisp->QueryInterface(IID_IHTMLDocument2, (void**)&pDoc);
				pDisp->Release();
				if (FAILED(hr))
				{
					mb(L"Retrieve IHTMLDoc2 gave an error", L"385suhfgdju");
					return E_UNEXPECTED;
				}

				OnDocumentComplete(pDoc);
				pDoc->Release();
			}
		}
		return S_OK;
	}

private:
	/* member variables */
	unordered_set<wstring> usetPageMWList;

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
	DWORD dwCookie, dwCookie2;
	BOOL bDocumentCompleted;
	long ref;
	mysqlpp::Connection conn;
	tiodbc::connection tConn;
	tiodbc::statement tStmt;
	string LgID;
	wstring wLgID;
	tstring tLgID;
	wstring WordChars;
	wstring SentDelims;
	wstring SentDelimExcepts;
	wstring wstrDict1;
	wstring wstrDict2;
	wstring wstrGoogleTrans;
	vector<wstring> chunks;
	TokenStruct tokens, tknSansBookmarks, tknCanonical;
	bool bWithSpaces;
	_bstr_t bstrUrl;
	LWTCache cache;
	// wregex
	wregex_iterator rend;
	wregex wWordSansBookmarksWrgx;
	wregex TagNotTagWrgx;
};

#endif
