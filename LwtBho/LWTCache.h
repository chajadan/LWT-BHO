#ifndef LWTCache_INCLUDED
#define LWTCache_INCLUDED

#include <unordered_map>
#include <string>
#include <assert.h>
#include "chajUtil.h"
#include "TermRecord.h"

typedef std::unordered_map<wstring,TermRecord>::const_iterator cache_cit;
typedef std::unordered_map<wstring,TermRecord>::iterator cache_it;

class LWTCache
{
public:
	LWTCache(wstring wLgID = L"", wstring wTblPfx = L"") : _cur(0), _wLgID(wLgID), _wTblPfx(wTblPfx)
	{
		if (!InitializeCriticalSectionAndSpinCount(&CS_AddCacheEntry, 0x00000400))
			TRACE(L"Cound not initialize critical section in LWTCache. 3850duhf\n");
	}
	~LWTCache()
	{
		DeleteCriticalSection(&CS_AddCacheEntry);
	}
	wstring getLgID() {return _wLgID;}
	void erase(std::unordered_map<std::wstring,TermRecord>::iterator eraseIter)
	{
		_cache.erase(eraseIter);
	}
	std::unordered_map<std::wstring,TermRecord>::iterator find(const std::wstring& entry) 
	{
		std::unordered_map<std::wstring,TermRecord>::iterator it;
		EnterCriticalSection(&CS_AddCacheEntry);
		it = _cache.find(entry);
		LeaveCriticalSection(&CS_AddCacheEntry);
		return it;
	}
	std::unordered_map<std::wstring,TermRecord>::const_iterator cfind(const std::wstring& entry)
	{
		std::unordered_map<std::wstring,TermRecord>::iterator it;
		EnterCriticalSection(&CS_AddCacheEntry);
		it = _cache.find(entry);
		LeaveCriticalSection(&CS_AddCacheEntry);
		return it;
	}
	std::unordered_map<std::wstring,TermRecord>::iterator end() {return _cache.end();}
	std::unordered_map<std::wstring,TermRecord>::const_iterator cend() {return _cache.cend();}
	void insert(std::unordered_map<std::wstring,TermRecord>::value_type newEntry)
	{
		EnterCriticalSection(&CS_AddCacheEntry);
		newEntry.second.uIdent = _cur++;
		_cache.insert(newEntry);
#ifdef _DEBUG
			cache_it it = _cache.find(newEntry.first);
			assert(it != _cache.end());
#endif
		LeaveCriticalSection(&CS_AddCacheEntry);
	}
	unsigned int insert(wstring wEntry)
	{
		TermRecord rec(L"-1");
		std::unordered_map<std::wstring,TermRecord>::value_type newEntry(wEntry, rec);
		EnterCriticalSection(&CS_AddCacheEntry);
		newEntry.second.uIdent = _cur++;
		_cache.insert(newEntry);
#ifdef _DEBUG
			cache_it it = _cache.find(newEntry.first);
			assert(it != _cache.end());
#endif
		LeaveCriticalSection(&CS_AddCacheEntry);
	}
private:
	LWTCache(const LWTCache&) {}
	LWTCache& operator=(const LWTCache&) {}
	CRITICAL_SECTION CS_AddCacheEntry;
	std::unordered_map<std::wstring,TermRecord> _cache;
	unsigned int _cur;
	wstring _wLgID;
	wstring _wTblPfx;
};

#endif