// Calling CoInitializeEx repeated must be balanced with CoUninitialize.
//
// Output:
//
// CoInitializeEx -> 00000000
// CoInitializeEx -> 00000001
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
	if (FAILED(hr))
	{
		CoUninitialize();
		return hr;
	}

	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	dbg << "CoInitializeEx -> " << hr << "\n"; // S_FALSE
	if (FAILED(hr))
	{
		CoUninitialize();
		CoUninitialize();
		return hr;
	}

	CoUninitialize();
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