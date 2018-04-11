// Calling CoInitializeEx / CoUninitialize
//  1. Recommended to be done once per thread
//  2. Initialize the COM environment, else, we cannot create COM objects
//  3. CoUninitialize destroys the COM environment
//  4. CoUninitialize must be balanced
//
// Output:
//
// CoInitializeEx -> 00000000
// Run -> 00000001

#include "stdafx.h"

//----------------------------------------------------------------------

class CDbg
{
public:
	void out(const char* fmt, ...) const
	{
		va_list args;
		va_start(args, fmt);
		char text[1024] = {};
		vsprintf_s(text, fmt, args);
		va_end(args);
		OutputDebugStringA(text);
	}
	CDbg& operator << (const char* str) { OutputDebugStringA(str); return *this; }
	CDbg& operator << (HRESULT hr) { out("%p", hr); return *this; }
};

CDbg dbg;

//----------------------------------------------------------------------

HRESULT Run()
{
	HRESULT hr = S_OK;

	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	dbg << "CoInitializeEx -> " << hr << "\n"; // S_OK
	if (FAILED(hr)) return hr;

	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	dbg << "CoInitializeEx -> " << hr << "\n"; // S_FALSE
	if (FAILED(hr)) return hr;

	CoUninitialize();
	CoUninitialize();

	return hr;
}

//----------------------------------------------------------------------

int main()
{
	HRESULT hr = Run();
	dbg << "Run -> " << hr << "\n";
	return 0;
}