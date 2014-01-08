#pragma once

#include "Chunc.h"
#include <list>
#include <string>
#include <memory>
#include <condition_variable>
#include <Windows.h>
#include <mshtml.h>
#include <thread>
#include "chajUtil.h"

class LwtBho;

class Chunc_CachePageWords : public Chunc
{
public:
	Chunc_CachePageWords(
		LwtBho* caller,
		LPSTREAM pDocStream,
		std::shared_ptr<std::vector<std::wstring>> sp_vWords,
		std::condition_variable*  p_cv,
		bool* pbHeadStartComplete)
		:
		mCaller(caller),
		mpDocStream(pDocStream),
		msp_vWords(sp_vWords),
		mp_cv(p_cv),
		mpbHeadStartComplete(pbHeadStartComplete),
		mi(0),
		msp_vWords_key(0),
		mi_key(0)
	{
		SetVariableBounds();
		AddChangeTriggers();
	}
	void operator()();
private:
	template <typename T>
	void AddChangeTrigger(ChuncVar<T>& var, std::function<bool(T)> condition, void(Chunc_CachePageWords::*Func)(void))
	{
		++mCurKey;
		std::function<bool(T)>* here = new std::function<bool(T)>;
		*here = condition;
		mTriggers.push_back(std::make_tuple(mCurKey, (void*)(here), Func));
		ListenToChanges<T>(var, mCurKey);
	}
	void HearChunkVarChange(int key, void* value)
	{
		for (auto& trigger : mTriggers)
		{
			if (key == std::get<0>(trigger))
			{
				void* condition = std::get<1>(trigger);
				std::function<bool(std::shared_ptr<std::vector<std::wstring>>)>* hereCond 
					= (std::function<bool(std::shared_ptr<std::vector<std::wstring>>)>*) condition;
				ChuncVar<std::shared_ptr<std::vector<std::wstring>>> hereVal
					= *((std::shared_ptr<std::vector<std::wstring>>*)(value));
				if ((*hereCond)(hereVal))
				{
					void(Chunc_CachePageWords::*Func)(void) = std::get<2>(trigger);
					(this->*Func)();
				}
			}
		}
	}
	template <typename T>
	void ListenToChanges(ChuncVar<T>& var, int key)
	{
		var.AddChangeListener(this, key);
	}

	ChuncVar<LPSTREAM> mpDocStream;
	ChuncVar<std::shared_ptr<std::vector<std::wstring>>> msp_vWords;
	ChuncVar<std::condition_variable*> mp_cv;
	ChuncVar<bool*> mpbHeadStartComplete;
	ChuncVar<int> mi;

	void NotifyHeadStart()
	{
		static bool bHasRun;
		if (bHasRun != true)
		{
			bHasRun = true;
			{
				std::mutex m;
				std::lock_guard<std::mutex> lk(m);
				*((bool*)mpbHeadStartComplete) = true;
				// also simply *mpbHeadStartComplete = true seems to work
			}
			((std::condition_variable*)mp_cv)->notify_all();
		}
	}

	void SetVariableBounds()
	{
		mpDocStream.AddBound([](LPSTREAM x){return !x;});
		msp_vWords.AddBound([](std::shared_ptr<std::vector<std::wstring>> x){return !x;});
	}

	void AddChangeTriggers()
	{
		msp_vWords_key = mCurKey + 1;
		AddChangeTrigger<std::shared_ptr<std::vector<std::wstring>>>(msp_vWords, [](std::shared_ptr<std::vector<std::wstring>> x){return (*(std::shared_ptr<std::vector<std::wstring>>)x).size() == 0;}, &Chunc_CachePageWords::NotifyHeadStart);

		mi_key = mCurKey + 1;
		AddChangeTrigger<int>(mi, [](int x){return x == 1;}, &Chunc_CachePageWords::NotifyHeadStart);
	}

	static const int MAX_MYSQL_INLIST_LEN = 500;
	int msp_vWords_key;
	int mi_key;
	std::vector<std::tuple<int, void*, void (Chunc_CachePageWords::*)(void)>> mTriggers;
	LwtBho* mCaller;
};