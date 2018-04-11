//----------------------------------------------------------------------
// Calling CoInitializeEx repeated must be balanced with CoUninitialize.
//
// Output (2 CoInitializeEx version):
//
// main::Run Begin
// CoInitializeEx Begin
// 'IntroToComV2.exe' (Win32) : Loaded 'C:\Windows\System32\kernel.appcore.dll'.Cannot find or open the PDB file.
// 'IntroToComV2.exe' (Win32) : Loaded 'C:\Windows\System32\msvcrt.dll'.Cannot find or open the PDB file.
// CoInitializeEx End 0x00000000 (The operation completed successfully.)
// CoInitializeEx Begin
// CoInitializeEx End 0x00000001 (Incorrect function.)
// CoInitializeEx Begin
// CoUninitialize End
// CoInitializeEx Begin
// CoUninitialize End
// main::Run End hr : 0x00000001 (Incorrect function.)
//
// Output (3 CoInitializeEx version):
//
// main::Run Begin
// CoInitializeEx Begin
// 'IntroToComV2.exe' (Win32) : Loaded 'C:\Windows\System32\kernel.appcore.dll'.Cannot find or open the PDB file.
// 'IntroToComV2.exe' (Win32) : Loaded 'C:\Windows\System32\msvcrt.dll'.Cannot find or open the PDB file.
// CoInitializeEx End 0x00000000 (The operation completed successfully.)
// CoInitializeEx Begin
// CoInitializeEx End 0x00000001 (Incorrect function.)
// CoInitializeEx Begin
// CoInitializeEx End 0x80010106 (Cannot change thread mode after it is set.)
// CoInitializeEx Begin
// CoUninitialize End
// CoInitializeEx Begin
// CoUninitialize End
// main::Run End 0x80010106 (Cannot change thread mode after it is set.)
//----------------------------------------------------------------------

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

	dbg << "CoInitializeEx Begin\n";
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	Sleep(200);
	dbg << "CoInitializeEx End " << hr << "\n"; // S_OK - 0x00000000 (The operation completed successfully.)
	if (FAILED(hr)) return hr;

	dbg << "CoInitializeEx Begin\n";
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	Sleep(200);
	dbg << "CoInitializeEx End " << hr << "\n"; // S_FALSE - 0x00000001 (Incorrect function.)
	if (FAILED(hr))
	{
		dbg << "CoInitializeEx Begin\n";
		CoUninitialize();
		Sleep(200);
		dbg << "CoUninitialize End\n";
		return hr;
	}

	//dbg << "CoInitializeEx Begin\n";
	//hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	//Sleep(200);
	//dbg << "CoInitializeEx End " << hr << "\n"; // RPC_E_CHANGED_MODE - 0x80010106 (Cannot change thread mode after it is set.)
	//if (FAILED(hr))
	//{
	//	dbg << "CoInitializeEx Begin\n";
	//	CoUninitialize();
	//	Sleep(200);
	//	dbg << "CoUninitialize End\n";
	//	dbg << "CoInitializeEx Begin\n";
	//	CoUninitialize();
	//	Sleep(200);
	//	dbg << "CoUninitialize End\n";
	//	return hr;
	//}

	//dbg << "CoInitializeEx Begin\n";
	//CoUninitialize();
	//Sleep(200);
	//dbg << "CoUninitialize End\n";

	dbg << "CoInitializeEx Begin\n";
	CoUninitialize();
	Sleep(200);
	dbg << "CoUninitialize End\n";
	dbg << "CoInitializeEx Begin\n";
	CoUninitialize();
	Sleep(200);
	dbg << "CoUninitialize End\n";

	return hr;
}

//----------------------------------------------------------------------

int main()
{
	dbg << "main::Run Begin\n";
	HRESULT hr = Run();
	dbg << "main::Run End " << hr << "\n";
	return 0;
}