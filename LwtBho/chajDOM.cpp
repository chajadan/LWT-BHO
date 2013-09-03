#include "chajDOM.h"
#include "chajUtil.h"
#include "comip.h"

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
IHTMLElement* chaj::DOM::GetHTMLElementFromDispatch(IDispatch* pDispElement)
{
	IHTMLElement* res = NULL;
	pDispElement->QueryInterface(IID_IHTMLElement, (void**)&res);
	return res;
}

IDOMTreeWalker* chaj::DOM::GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow)
{
	chaj::DOM::DOMIteratorFilter filter(FilterFunc);
	return GetTreeWalkerWithFilter(pDoc, pRootAt, &filter, lShow);
}

IDOMTreeWalker* chaj::DOM::GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* pFilter, long lShow)
{
	IDocumentTraversal* pDT =  GetDocTravFromDoc(pDoc);
	if (pDT == nullptr)
		return nullptr;

	IDOMTreeWalker* pTreeWalker = nullptr;
	VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = pFilter;
	pDT->createTreeWalker(pRootAt, lShow, &varFilter, VARIANT_TRUE, &pTreeWalker);
	pDT->Release();
	return pTreeWalker;
}

IDOMNodeIterator* chaj::DOM::GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow)
{
	DOMIteratorFilter filter(FilterFunc);
	return GetNodeIteratorWithFilter(pDoc, pRootAt, &filter, lShow);
}
IDOMNodeIterator* chaj::DOM::GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* pFilter, long lShow)
{
	IDocumentTraversal* pDT =  GetDocTravFromDoc(pDoc);
	if (pDT == nullptr)
		return nullptr;

	IDOMNodeIterator* pNodeIterator = nullptr;
	VARIANT varFilter; VariantInit(&varFilter); varFilter.vt = VT_DISPATCH; varFilter.pdispVal = pFilter;
	pDT->createNodeIterator(pRootAt, lShow, &varFilter, VARIANT_TRUE, &pNodeIterator);
	pDT->Release();
	return pNodeIterator;
}
IDocumentTraversal* chaj::DOM::GetDocTravFromDoc(IHTMLDocument2* pDoc)
{
	IDocumentTraversal* res = nullptr;
	HRESULT hr = pDoc->QueryInterface(IID_IDocumentTraversal, reinterpret_cast<void**>(&res));
	return res;
}
IHTMLBodyElement* chaj::DOM::GetBodyElementFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLBodyElement* res = nullptr;
	IHTMLElement* pBody = GetBodyAsElementFromDoc(pDoc);
	pBody->QueryInterface(IID_IHTMLBodyElement, reinterpret_cast<void**>(&res));
	return res;
}
IHTMLElement* chaj::DOM::GetBodyAsElementFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLElement* res = nullptr;
	pDoc->get_body(&res);
	return res;
}

IHTMLTxtRange* chaj::DOM::GetBodyTxtRangeFromDoc(IHTMLDocument2* pDoc)
{
	IHTMLBodyElement* pBodyElement = GetBodyElementFromDoc(pDoc);
	IHTMLTxtRange* pTxtRange = nullptr;
	pBodyElement->createTextRange(&pTxtRange);
	return pTxtRange;
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
	return iEC;
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
	if (!iDisp)
		return nullptr;
	else
	{
		IHTMLElement* retVal = chaj::COM::GetAlternateInterface<IDispatch,IHTMLElement>(iDisp);
		iDisp->Release();
		return retVal;
	}
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
	HRESULT hr = pElement->get_className(&bstrStat);
	std::wstring out(L"");
	if (bstrStat != NULL)
	{
		out.append(bstrStat);
		SysFreeString(bstrStat);
	}
	return out;
}
IHTMLElement* chaj::DOM::GetHeadAsElementFromDoc(IHTMLDocument2* pDoc)
{
	if (!pDoc)
		return nullptr;
	IHTMLElementCollection* iEC = nullptr;
	HRESULT hr = pDoc->get_all(&iEC);
	if (FAILED(hr))
		return nullptr;
	chaj::COM::SmartCOMRelease srEC(iEC);

	BSTR bstrHead = SysAllocString(L"HEAD");
	VARIANT varName; VariantInit(&varName); varName.vt = VT_BSTR; varName.bstrVal = bstrHead;
	VARIANT varIndex; VariantInit(&varIndex); varIndex.vt = VT_I4; varIndex.lVal = 0;
	IDispatch* pDispHead = nullptr;
	hr = iEC->item(varName, varIndex, &pDispHead);
	if (FAILED(hr) || !pDispHead)
	{
		iEC->tags(varName, &pDispHead);
		if (!pDispHead)
		{
			VariantClear(&varName);
			VariantClear(&varIndex);
			return nullptr;
		}
		IHTMLElementCollection* iEC2 = chaj::COM::GetAlternateInterface<IDispatch,IHTMLElementCollection>(pDispHead);
		pDispHead->Release();
		iEC2->item(varIndex, varIndex, &pDispHead);
		iEC2->Release();
		if (!pDispHead)
		{
			VariantClear(&varName);
			VariantClear(&varIndex);
			return nullptr;
		}
	}
	VariantClear(&varIndex);
	VariantClear(&varName);
	chaj::COM::SmartCOMRelease srDispHead(pDispHead);

	return chaj::COM::GetAlternateInterface<IDispatch,IHTMLElement>(pDispHead);
}
HRESULT chaj::DOM::TxtRange_CollapseToEnd(IHTMLTxtRange* pRange)
{
	_bstr_t bstrHowEndToStart(L"StartToEnd");
	HRESULT hr = pRange->setEndPoint(bstrHowEndToStart.GetBSTR(), pRange);
	return hr;
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