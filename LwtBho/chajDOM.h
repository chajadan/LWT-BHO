#ifndef chajDOM_INCLUDED
#define chajDOM_INCLUDED

#include <Windows.h>	// IDispatch
#include <string>		// std::wstring
#include "mshtml.h"		// IHTMLElement, etc.
#include "comutil.h"	// _bstr_t(Lib: comsuppw.lib or comsuppwd.lib)

namespace chaj
{
	namespace DOM
	{
		extern const long FILTER_ACCEPT;
		extern const long FILTER_REJECT;
		extern const long FILTER_SKIP;

		extern const long SHOW_ALL;
		extern const long SHOW_ELEMENT;
		
		std::wstring GetElementClass(IHTMLElement* pElement);
		HRESULT SetElementClass(IHTMLElement* pElement, const std::wstring& wstrNewClass);
		IHTMLElement* GetElementFromId(const std::wstring& wstrId, IHTMLDocument2* pDoc);
		IHTMLElementCollection* GetAllElementsFromDoc(IHTMLDocument2* pDoc);
		IHTMLElement* GetHeadAsElementFromDoc(IHTMLDocument2* pDoc);
		IDocumentTraversal* GetDocTravFromDoc(IHTMLDocument2* pDoc);
		IDOMTreeWalker* GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* pFilter, long lShow = SHOW_ALL);
		IDOMTreeWalker* GetTreeWalkerWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow = SHOW_ALL);
		IDOMNodeIterator* GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, long (*FilterFunc)(IDispatch*), long lShow = SHOW_ALL);
		IDOMNodeIterator* GetNodeIteratorWithFilter(IHTMLDocument2* pDoc, IDispatch* pRootAt, IDispatch* pFilter, long lShow = SHOW_ALL);

		HRESULT AppendStylesToDoc(const std::wstring& styles, IHTMLDocument2* pDoc);
		HRESULT AppendHTMLAfterBegin(const std::wstring& toAppend, IHTMLElement* pElement);
		HRESULT AppendHTMLBeforeBegin(const std::wstring& toAppend, IHTMLElement* pElement);
		HRESULT AppendHTMLBeforeEnd(const std::wstring& toAppend, IHTMLElement* pElement);
		HRESULT AppendHTMLAdjacent(const std::wstring& wWhere, const std::wstring& toAppend, IHTMLElement* pElement);

		HRESULT TextRange_SetEndPoint(_bstr_t& how, IHTMLTxtRange* pRange, IHTMLTxtRange* pRefRange);
		HRESULT TxtRange_CollapseToBegin(IHTMLTxtRange* pRange);
		HRESULT TxtRange_CollapseToEnd(IHTMLTxtRange* pRange);
		HRESULT TxtRange_RevertEnd(IHTMLTxtRange* pRange);
		HRESULT TxtRange_MoveStartByChars(IHTMLTxtRange* pRange, long lngChars);

		inline std::wstring HTMLTxtRange_get_text(IHTMLTxtRange* pRange)
		{
			std::wstring wstrText(L"");
			
			if (pRange)
			{
				BSTR bstrText = nullptr;
				pRange->get_text(&bstrText);

				if (bstrText)
				{
					wstrText += bstrText;
					SysFreeString(bstrText);
				}
			}

			return wstrText;
		}

		inline std::wstring HTMLTxtRange_htmlText(IHTMLTxtRange* pRange)
		{
			std::wstring retVal(L"");

			if (pRange)
			{
				BSTR bstrHTMLText = nullptr;
				pRange->get_htmlText(&bstrHTMLText);

				if (bstrHTMLText)
				{
					retVal += bstrHTMLText;
					SysFreeString(bstrHTMLText);
				}
			}

			return retVal;
		}

		class DOMIteratorFilter: public IDispatch
		{
		public:
			DOMIteratorFilter(long (*FilterFunc)(IDispatch*)) : ff(FilterFunc) {}

			ULONG STDMETHODCALLTYPE AddRef();
			ULONG STDMETHODCALLTYPE Release();
			HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);

			HRESULT STDMETHODCALLTYPE GetTypeInfoCount(unsigned int FAR* pctinfo);
			HRESULT STDMETHODCALLTYPE GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR*  ppTInfo);
			HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId);
			HRESULT STDMETHODCALLTYPE Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

		private:
			long NodeFilterStatus(IDispatch* pNode);

		private:
			long ref;
			long (*ff)(IDispatch*);
		};

		//CIC /* {C}ommon {I}nterface {C}hains */
		IHTMLElement* GetBodyAsElementFromDoc(IHTMLDocument2* pDoc);
		IHTMLBodyElement* GetBodyElementFromDoc(IHTMLDocument2* pDoc);
		IHTMLTxtRange* GetBodyTxtRangeFromDoc(IHTMLDocument2* pDoc);

		std::wstring GetAttributeValue(IHTMLElement* pElement, const std::wstring& wstrAttribute, long lFlags = 0);
		HRESULT SetAttributeValue(IHTMLElement* pElement, const std::wstring& wstrAttribute, const std::wstring& wstrAttValue, long lFlags = 0);
		HRESULT SetElementInnerText(IHTMLElement* pElement, const std::wstring& wstrInnerText);


		IHTMLElement* GetHTMLElementFromDispatch(IDispatch* pDispElement);
	}
}

#endif