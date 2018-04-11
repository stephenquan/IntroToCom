// Calling CoInitializeEx / CoUninitialize
//  1. Recommended to be done once per thread
//  2. Initialize the COM environment, else, we cannot create COM objects
//  3. CoUninitialize destroys the COM environment
//  4. CoUninitialize must be balanced
//
// Output:
//
// main::Run Begin
// Run::CoInitializeEx Begin
// 'IntroToComV1.exe' (Win32) : Loaded 'C:\Windows\System32\kernel.appcore.dll'.Cannot find or open the PDB file.
// 'IntroToComV1.exe' (Win32) : Loaded 'C:\Windows\System32\msvcrt.dll'.Cannot find or open the PDB file.
// Run::CoInitializeEx End 0x00000000 (The operation completed successfully.)
// Run::CoUninitialize Begin
// Run::CoUninitialize End
// main::Run End 0x00000000 (The operation completed successfully.)

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
	void out(HRESULT hr)
	{
		CComBSTR bstrError;
		_com_error err(hr);
		out("0x%-8.8x (%S)", hr, err.ErrorMessage());
	}
	CDbg& operator << (const char* str) { OutputDebugStringA(str); return *this; }
	CDbg& operator << (HRESULT hr) { out(hr); return *this; }
	CDbg& operator << (int v) { out("%d", v); return *this; }
	CDbg& operator << (VARIANT_BOOL b) { return CDbg::operator<<(b == VARIANT_TRUE ? "VARIANT_TRUE" : "VARIANT_FALSE"); }
	CDbg& operator << (bool b) { return CDbg::operator<<(b ? "true" : "false"); }
	CDbg& operator << (ULONG v) { out("%d", v); return *this; }
};

CDbg dbg;

//----------------------------------------------------------------------

HRESULT Run()
{
	HRESULT hr = S_OK;

	dbg << "Run::CoInitializeEx Begin\n";
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	Sleep(200);
	dbg << "Run::CoInitializeEx End " << hr << "\n"; // S_OK
	if (FAILED(hr)) return hr;

	// ... Do COM stuff here ...

	dbg << "Run::CoUninitialize Begin\n";
	CoUninitialize();
	Sleep(200);
	dbg << "Run::CoUninitialize End\n";

	return hr;
}

//----------------------------------------------------------------------

int main()
{
	HRESULT hr = S_OK;
	
	dbg << "main::Run Begin\n";
	hr = Run();
	dbg << "main::Run End " << hr << "\n";

	return 0;
}