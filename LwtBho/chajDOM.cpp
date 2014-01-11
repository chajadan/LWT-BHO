#include "chajDOM.h"
#include "chajUtil.h"
#include "comip.h"

using namespace chaj::COM;

namespace chaj
{
	LONG gref=0;

	namespace DOM
	{
		extern const long FILTER_ACCEPT = 1;
		extern const long FILTER_REJECT = 2;
		extern const long FILTER_SKIP = 3;

		extern const long SHOW_ALL = 0xffffffff;
		extern const long SHOW_ELEMENT = 1;
		extern const long SHOW_ATTRIBUTE = 2;
		extern const long SHOW_TEXT = 4;

		extern const long NODE_TYPE_ELEMENT = 1;
		extern const long NODE_TYPE_TEXT = 3;
	}
}

ULONG STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::AddRef() {InterlockedIncrement(&gref); return InterlockedIncrement(&ref);}
ULONG STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::Release() {int tmp=InterlockedDecrement(&ref); if (tmp==0) delete this; InterlockedDecrement(&gref); return tmp;}
HRESULT STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::QueryInterface(REFIID riid, void **ppv)
{
	if (riid==IID_IUnknown)
		*ppv=static_cast<DOMIteratorFilter*>(this);
	else if (riid==IID_IDispatch)
		*ppv=static_cast<IDispatch*>(this);
	else
		return E_NOINTERFACE;
		
	AddRef();
	return S_OK;
}

// IDispatch implementation
HRESULT STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::GetTypeInfoCount(unsigned int FAR* pctinfo) {*pctinfo=1; return NOERROR;}
#pragma warning( push )
#pragma warning( disable : 4100 )
HRESULT STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR*  ppTInfo) {return NOERROR;}
HRESULT STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId) {return NOERROR;}
HRESULT STDMETHODCALLTYPE chaj::DOM::DOMIteratorFilter::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	if (dispidMember == 0)
	{
		if (wFlags == DISPATCH_METHOD)
		{
			(*pvarResult).vt = VT_I4;
			(*pvarResult).lVal = ff(pDispParams->rgvarg->pdispVal);
		}
	}
	return S_OK;
}
#pragma warning ( pop )

std::wstring chaj::DOM::GetAttributeValue(IHTMLElement* pElement, const std::wstring& wstrAttribute, long lFlags)
{
	_bstr_t bstrAttribute(wstrAttribute.c_str());
	VARIANT varAttValue; VariantInit(&varAttValue);
	HRESULT hr = pElement->getAttribute(bstrAttribute.GetBSTR(), lFlags, &varAttValue);
	if (FAILED(hr))
		return std::wstring(L"");
	else
	{
		std::wstring wRetVal(L"");
		if (varAttValue.vt == VT_BSTR && varAttValue.bstrVal != NULL)
			wRetVal = std::wstring(varAttValue.bstrVal);

		VariantClear(&varAttValue);
		return wRetVal;
	}
}
HRESULT chaj::DOM::SetAttributeValue(IHTMLElement* pElement, const std::wstring& wstrAttribute, const std::wstring& wstrAttValue, long lFlags)
{
	_bstr_t bstrAttribute(wstrAttribute.c_str());
	BSTR bstrAttValue = SysAllocString(wstrAttValue.c_str());
	VARIANT varAttValue; VariantInit(&varAttValue); varAttValue.vt = VT_BSTR; varAttValue.bstrVal = bstrAttValue;
	HRESULT hr = pElement->setAttribute(bstrAttribute.GetBSTR(), varAttValue, lFlags);
	VariantClear(&varAttValue);
	return hr;
}
HRESULT chaj::DOM::SetElementInnerText(IHTMLElement* pElement, const std::wstring& wstrInnerText)
{
	BSTR bstrOut = SysAllocString(wstrInnerText.c_str());
	HRESULT hr = pElement->put_innerText(bstrOut);
	SysFreeString(bstrOut);
	return hr;
}
IHTMLElement* chaj::DOM::CreateElement(IHTMLDocument2* pDoc, const std::wstring& tag)
{
	SmartCom<IHTMLElement> pElement = nullptr;
	BSTR bstrTag = SysAllocString(tag.c_str());
	pDoc->createElement(bstrTag, pElement);
	SysFreeString(bstrTag);
	return pElement.Relinquish();
}
HRESULT chaj::DOM::SetElementOuterHTML(IHTMLElement* pElement, const std::wstring& wstrOuterHTML)
{
	BSTR bstrOut = SysAllocString(wstrOuterHTML.c_str());
	HRESULT hr = pElement->put_outerHTML(bstrOut);
	SysFreeString(bstrOut);
	return hr;
}
IHTMLDOMTextNode* chaj::DOM::SplitTextNode(IHTMLDOMTextNode* pTextNode, long offset)
{
	SmartCom<IHTMLDOMNode> pNode;

	HRESULT hr = pTextNode->splitText(offset, pNode);
	if (SUCCEEDED(hr) && pNode)
		return GetAlternateInterface<IHTMLDOMNode,IHTMLDOMTextNode>(pNode);
	else
		return nullptr;
}
std::wstring chaj::DOM::GetTagFromElement(IHTMLElement* pElement)
{
	_bstr_t bstrTag;
	std::wstring wTag(L"");

	HRESULT hr = pElement->get_tagName(bstrTag.GetAddress());
	if (SUCCEEDED(hr) && bstrTag.length())
		wTag = bstrTag;

	return wTag;
}
IDOMTreeWalker* chaj::DOM::GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow)
{
	chaj::DOM::DOMIteratorFilter filter(FilterFunc);
	return GetTreeWalkerWithFilter(pDoc, pRootAt, &filter, lShow);
}

IDOMTreeWalker* chaj::DOM::GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* s_pFilter, long lShow)
{
	SmartCom<IDocumentTraversal> pDT =  GetDocTravFromDoc(pDoc);
	if (!pDT)
		return nullptr;

	SmartCom<IDOMTreeWalker> pTreeWalker;
	VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = s_pFilter;
	pDT->createTreeWalker(pRootAt, lShow, &varFilter, VARIANT_TRUE, pTreeWalker);
	return pTreeWalker.Relinquish();
}

IDOMNodeIterator* chaj::DOM::GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow)
{
	DOMIteratorFilter filter(FilterFunc);
	return GetNodeIteratorWithFilter(pDoc, pRootAt, &filter, lShow);
}
IDOMNodeIterator* chaj::DOM::GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* s_pFilter, long lShow)
{
	IDocumentTraversal* pDT =  GetDocTravFromDoc(pDoc);
	if (pDT == nullptr)
		return nullptr;

	IDOMNodeIterator* pNodeIterator = nullptr;
	VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = s_pFilter;
	pDT->createNodeIterator(pRootAt, lShow, &varFilter, VARIANT_TRUE, &pNodeIterator);
	pDT->Release();
	return pNodeIterator;
}
IDocumentTraversal* chaj::DOM::GetDocTravFromDoc(IHTMLDocument2* pDoc)
{
	IDocumentTraversal* res = nullptr;
	HRESULT hr = pDoc->QueryInterface(IID_IDocumentTraversal, reinterpret_cast<void**>(&res));
	if (FAILED(hr) || !res)
		res = nullptr;

	return res;
}
IHTMLBodyElement* chaj::DOM::GetBodyElementFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLBodyElement* res = nullptr;
	SmartCom<IHTMLElement> pBody = GetBodyFromDoc(pDoc);
	if (pBody)
	{
		pBody->QueryInterface(IID_IHTMLBodyElement, reinterpret_cast<void**>(&res));
		pBody->Release();
	}
	return res;
}
IHTMLElement* chaj::DOM::GetBodyFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLElement* res = nullptr;
	pDoc->get_body(&res);
	return res;
}

IHTMLTxtRange* chaj::DOM::GetBodyTxtRangeFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLTxtRange* pTxtRange = nullptr;
	IHTMLBodyElement* pBodyElement = GetBodyElementFromDoc(pDoc);
	if (pBodyElement)
	{
		pBodyElement->createTextRange(&pTxtRange);
		pBodyElement->Release();
	}
	return pTxtRange;
}
IHTMLElement* chaj::DOM::GetClickedElementFromEvent(IHTMLEventObj* pEvent)
{
	long x, y;
	pEvent->get_clientX(&x);
	pEvent->get_clientY(&y);

	SmartCom<IHTMLElement> pElement;
	HRESULT hr = pEvent->get_srcElement(pElement);
	if (FAILED(hr))
	{
		TRACE(L"%s", L"Leaving GetClickedElementFromEvent with error result\n");
		return nullptr;
	}

	return pElement.Relinquish();
}
HRESULT chaj::DOM::AppendStylesToDoc(const std::wstring& styles, IHTMLDocument2* pDoc)
{
	if (!pDoc)
		return E_INVALIDARG;

	BSTR bstrURL = SysAllocString(L"");
	IHTMLStyleSheet* pStyle = nullptr;
	HRESULT hr = pDoc->createStyleSheet(bstrURL, 0, &pStyle);
	SysFreeString(bstrURL);
	if (FAILED(hr))
		return hr;

	BSTR bstrStyles = SysAllocString(styles.c_str());
	hr = pStyle->put_cssText(bstrStyles);
	SysFreeString(bstrStyles);
	pStyle->Release();
	
	return hr;
}
HRESULT chaj::DOM::AppendHTMLBeforeBegin(const std::wstring& toAppend, IHTMLElement* pElement)
{
	return chaj::DOM::AppendHTMLAdjacent(L"beforeBegin", toAppend, pElement);
}
HRESULT chaj::DOM::AppendHTMLBeforeEnd(const std::wstring& toAppend, IHTMLElement* pElement)
{
	return chaj::DOM::AppendHTMLAdjacent(L"beforeEnd", toAppend, pElement);
}
HRESULT chaj::DOM::AppendHTMLAfterBegin(const std::wstring& toAppend, IHTMLElement* pElement)
{
	return chaj::DOM::AppendHTMLAdjacent(L"afterBegin", toAppend, pElement);
}
HRESULT chaj::DOM::AppendHTMLAdjacent(const std::wstring& wWhere, const std::wstring& toAppend, IHTMLElement* pElement)
{
	BSTR bstrWhere = SysAllocString(wWhere.c_str());
	BSTR bstrToAppend = SysAllocString(toAppend.c_str());
	HRESULT hr = pElement->insertAdjacentHTML(bstrWhere, bstrToAppend);
	SysFreeString(bstrToAppend);
	SysFreeString(bstrWhere);
	return hr;
}
IHTMLElementCollection* chaj::DOM::GetAllElementsFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLElementCollection* iEC = nullptr;
	HRESULT hr = pDoc->get_all(&iEC);
	if (FAILED(hr))
		iEC = nullptr;
	return iEC;
}
IHTMLElement* chaj::DOM::GetElementFromDisp(IDispatch* pDispElement)
{
	IHTMLElement* res = NULL;
	pDispElement->QueryInterface(IID_IHTMLElement, (void**)&res);
	return res;
}
IHTMLElement* chaj::DOM::GetElementFromId(const std::wstring& wstrId, IHTMLDocument2* pDoc)
{
	IHTMLElementCollection* iEC = GetAllElementsFromDoc(pDoc);
	if (!iEC)
		return nullptr;

	IDispatch* iDisp = nullptr;
	BSTR bstrId = SysAllocString(wstrId.c_str());
	VARIANT vId; VariantInit(&vId); vId.vt = VT_BSTR; vId.bstrVal = bstrId;
	VARIANT vIndex; VariantInit(&vIndex); vIndex.vt = VT_I4; vIndex.lVal = 0;
	HRESULT hr = iEC->item(vId, vIndex, &iDisp);
	iEC->Release();
	VariantClear(&vId);
	VariantClear(&vIndex);
	if (FAILED(hr) || !iDisp)
		return nullptr;
	else
	{
		IHTMLElement* retVal = chaj::COM::GetAlternateInterface<IDispatch,IHTMLElement>(iDisp);
		iDisp->Release();
		return retVal;
	}
}
IHTMLEventObj* chaj::DOM::GetEventFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLWindow2* window = NULL;
	HRESULT hr = pDoc->get_parentWindow(&window);
	if (FAILED(hr))
	{
		TRACE(L"%s", L"Leaving GetEventFromDoc with error result\n");
		return nullptr;
	}

	IHTMLEventObj* event;
	hr = window->get_event(&event);
	window->Release();
	if (FAILED(hr))
	{
		TRACE(L"%s", L"Leaving GetEventFromDoc with error result\n");
		return nullptr;
	}

	return event;
}
HRESULT chaj::DOM::SetElementClass(IHTMLElement* pElement, const std::wstring& wstrNewClass)
{
	BSTR bstrNewClass = SysAllocString(wstrNewClass.c_str());
	HRESULT hr = pElement->put_className(bstrNewClass);
	SysFreeString(bstrNewClass);
	return hr;
}
std::wstring chaj::DOM::GetElementClass(IHTMLElement* pElement)
{
	BSTR bstrStat = NULL;
	std::wstring out;
	HRESULT hr = pElement->get_className(&bstrStat);
	if (!FAILED(hr) && bstrStat != NULL)
	{
		out.append(bstrStat);
		SysFreeString(bstrStat);
	}
	return out;
}
IHTMLElement* chaj::DOM::GetHeadFromDoc(IHTMLDocument2* pDoc)
{
	if (!pDoc)
		return nullptr;
	SmartCom<IHTMLElementCollection> iEC;
	HRESULT hr = pDoc->get_all(iEC);
	if (FAILED(hr) || !iEC)
		return nullptr;

	BSTR bstrHead = SysAllocString(L"HEAD");
	VARIANT varName; VariantInit(&varName); varName.vt = VT_BSTR; varName.bstrVal = bstrHead;
	VARIANT varIndex; VariantInit(&varIndex); varIndex.vt = VT_I4; varIndex.lVal = 0;
	SmartCom<IDispatch> pDispHead;
	hr = iEC->item(varName, varIndex, pDispHead);
	if (FAILED(hr) || !pDispHead)
	{
		iEC->tags(varName, pDispHead);
		if (!pDispHead)
		{
			VariantClear(&varName);
			VariantClear(&varIndex);
			return nullptr;
		}
		SmartCom<IHTMLElementCollection> iEC2 = GetAlternateInterface<IDispatch,IHTMLElementCollection>(pDispHead);
		iEC2->item(varIndex, varIndex, pDispHead);
		if (!pDispHead)
		{
			VariantClear(&varName);
			VariantClear(&varIndex);
			return nullptr;
		}
	}
	VariantClear(&varIndex);
	VariantClear(&varName);

	return GetAlternateInterface<IDispatch,IHTMLElement>(pDispHead);
}
HRESULT chaj::DOM::TxtRange_CollapseToEnd(IHTMLTxtRange* pRange)
{
	_bstr_t bstrHowEndToStart(L"StartToEnd");
	return TextRange_SetEndPoint(bstrHowEndToStart, pRange, pRange);
}
HRESULT chaj::DOM::TxtRange_CollapseToBegin(IHTMLTxtRange* pRange)
{
	_bstr_t bstrHowEndToStart(L"EndToStart");
	return TextRange_SetEndPoint(bstrHowEndToStart, pRange, pRange);
}
HRESULT chaj::DOM::TextRange_SetEndPoint(_bstr_t& how, IHTMLTxtRange* pRange, IHTMLTxtRange* pRefRange)
{
	return pRange->setEndPoint(how.GetBSTR(), pRefRange);
}
HRESULT chaj::DOM::TxtRange_RevertEnd(IHTMLTxtRange* pRange)
{
	_bstr_t bstrUnit(L"textedit");
	long lngActual;
	HRESULT hr = pRange->moveEnd(bstrUnit.GetBSTR(), 1, &lngActual);
	return hr;
}
HRESULT chaj::DOM::TxtRange_MoveStartByChars(IHTMLTxtRange* pRange, long lngChars)
{
	_bstr_t bstrUnit(L"character");
	long lngActual;
	HRESULT hr = pRange->moveStart(bstrUnit.GetBSTR(), lngChars, &lngActual);
	return hr;

}

IHTMLDocument2* chaj::DOM::GetDocumentFromBrowser(IWebBrowser2* pBrowser)
{
	SmartCom<IDispatch> pDisp;
	HRESULT hr = pBrowser->get_Document(pDisp);
	if(FAILED(hr) || !pDisp)
	{
		TRACE(L"err: 4645sihdf\n");
		return nullptr;
	}

	return GetAlternateInterface<IDispatch,IHTMLDocument2>(pDisp);
}