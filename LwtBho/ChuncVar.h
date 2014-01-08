#pragma once

#include "Chunc.h"
#include "ChuncException.h"
#include <functional>
class Chunc;

class ChuncVarException : public ChuncException
{
};

class ChuncVar_OutOfBoundsException : public ChuncVarException
{
};


template <typename T>
class ChuncVar
{
public:
	ChuncVar(T value)
	{
		this->operator=(value);
	}
	void AddBound(std::function<bool(T)> bound, bool bCheckCurrentState = true)
	{
		mBounds.push_back(bound);
	}
	void AddChangeListener(Chunc* listener, int key, bool bNotifyCurrentState = true, bool bAllowDuplicate = false)
	{
		mChangeListeners.push_back(std::make_pair(listener, key));
		NotifyChangeListener(listener, key);
	}
	T& Get()
	{
		OnAccess();
		return mVar;
	}
	ChuncVar& operator=(T value)
	{
		ChangeValue(value);
		return *this;
	}
	ChuncVar& operator+=(T value)
	{
		mVar += value;
		return *this;
	}
	operator T()
	{
		return Get();
	}
protected:
	void OnAccess()
	{
	}
	void OnValueChange()
	{
		NotifyChangeListeners();
		CheckBounds();
	}
	/**
	 * ChangeValue()
	 * The sole function allowed to change the ChuncVar's value
	 */
	void ChangeValue(T value)
	{
		mVar = value;
		OnValueChange();
	}
	void CheckBounds()
	{
		for (auto& bound : mBounds)
			if (!bound(mVar))
				throw new ChuncVar_OutOfBoundsException();
	}
	void NotifyChangeListener(Chunc* listener, int key)
	{
		listener->HearChunkVarChange(key,  (void*)this);
	}
	void NotifyChangeListeners()
	{
		for (auto& entry : mChangeListeners)
			NotifyChangeListener(entry.first, entry.second);
	}
private:	
	T mVar;
	std::vector<std::function<bool(T)>> mBounds;
	std::vector<std::pair<Chunc*, int>> mChangeListeners;
};