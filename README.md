# IntroToCom

- IntroToComV1
  - Minimal app, demonstrating CoInitializeEx / CoUninitializeEx
  
- IntroToComV2
  - Nested CoInitializeEx / CoUninitializeEx
  
- IntroToComV3
  - XMLTest - Create Microsoft.XMLDOM object using smartpointers (two step)
  - XMLTest2 - Create Microsoft.XMLDOM object using smartpointers (single step)
    - Confidence means we can shorten the code (no need to troubleshoot)
  - XMLTest3 - Create Microsoft.XMLDOM object without smartpointers
    - 
    
Common SUCCESS codes:

| Name             | HRESULT    | Windows Description               | Explanation               |
| ---------------- | ---------- | --------------------------------- | ------------------------- |
| S_OK             | 0x00000000 |                                   |                           |
| DB_S_ENDOFROWSET | 0x00040EC6 |                                   | Last OLEDB record reached |
| S_FALSE          | 0x00000001 | Incorrect function.               | Cannot open file          |

Common Failure codes:

| Name                | HRESULT    | Windows Description               | Explanation                           |
| ------------------- | ---------- | --------------------------------- | ------------------------------------- |
| E_NOTIMPL           | 0x80004001 | Not implemented                   | Usually for Under Construction        |
| E_NOINTERFACE       | 0x80004002 | No such interface supported       | QueryInterface failed                 |
| E_FAIL              | 0x80040005 | Unspecified error                 | Overused exception code               |
| E_POINTER           | 0x80004003 | Invalid pointer                   | Typically used for invalid members    |
| E_UNEXPECTED        | 0X8000ffff | Catastrophic failure              | Unexpected edge condition hit         |
| CO_E_NOTINITIALIZED | 0x800401f0 | CoInitialize has not been called. | Self explanatory                      |
| CO_E_CLASSSTRING    | 0x800401f3 | Invalid class string              | Mispelt, not installed, not available |
| E_INVALIDARG        | 0x80070057 | One or more arguments are invalid | Self explanatory                      |
| E_OUTOFMEMORY       | 0x8007000E | Ran out of memory                 | Check for leaking memory patterns     |

Converting from DWORD to HRESULT:

    DWORD lastError = GetLastError();
    HRESULT hr = HRESULT_FROM_WIN(lastError)

See also

 - https://msdn.microsoft.com/en-us/library/windows/desktop/dd542643(v=vs.85).aspx
 - https://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx
