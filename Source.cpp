#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>

#include <iostream>
#include <regex>
#include <locale>

TCHAR szClassName[] = TEXT("Window");

void Replace(HWND hWnd, LPCTSTR lpszFilePath, LPCWSTR lpszFindText, LPCWSTR lpszToText)
{
	HANDLE hFile = CreateFile(lpszFilePath, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		DWORD dwFileSize = GetFileSize(hFile, 0);
		LPSTR lpszText = (LPSTR)GlobalAlloc(GPTR, dwFileSize + sizeof(WCHAR) * 16); // 16 文字分は予備
		DWORD dwReadSize;
		ReadFile(hFile, lpszText, dwFileSize, &dwReadSize, 0);
		int nCode = 0;
		// 0 : Shift-JIS
		// 1 : Unicode
		// 2 : UTF-8
		// まずはBOMで判定
		LPWSTR lpszNewStringW = 0;
		if (lpszText[0] == 0xEF && lpszText[1] == 0xBB && lpszText[2] == 0xBF)
		{
			nCode = 2;
			const int nLength = MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, 0, 0);
			lpszNewStringW = (LPWSTR)GlobalAlloc(GPTR, nLength * sizeof(WCHAR));
			MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, lpszNewStringW, nLength);
		}
		else if (lpszText[0] == 0xFF && lpszText[1] == 0xFE)
		{
			nCode = 1;
			LPWSTR lpszUnicode = (LPWSTR)lpszText[2];
			lpszNewStringW = (LPWSTR)GlobalAlloc(GPTR, lstrlenW(lpszUnicode) * sizeof(WCHAR));
			lstrcpyW(lpszNewStringW, lpszUnicode);
		}
		if (nCode == 0)
		{
			// Shif-JISかどうか判定
			int nLength = MultiByteToWideChar(932, 0, lpszText, -1, 0, 0);
			lpszNewStringW = (LPWSTR)GlobalAlloc(GPTR, nLength * sizeof(WCHAR));
			MultiByteToWideChar(932, 0, lpszText, -1, lpszNewStringW, nLength);
			nLength = WideCharToMultiByte(932, 0, lpszNewStringW, -1, 0, 0, 0, 0);
			LPSTR lpszStringA = (LPSTR)GlobalAlloc(GPTR, nLength);
			WideCharToMultiByte(932, 0, lpszNewStringW, -1, lpszStringA, nLength, 0, 0);
			const BOOL bShiftJIS = (lstrcmpA(lpszText, lpszStringA) == 0);
			GlobalFree(lpszStringA);
			if (bShiftJIS)
			{
				// Shif-JIS
				goto END_CODE_JUDGE;
			}
			else
			{
				// その他は Unicode と断定
				GlobalFree(lpszNewStringW);
				lpszNewStringW = (LPWSTR)GlobalAlloc(GPTR, lstrlenW((LPWSTR)lpszText) * sizeof(WCHAR));
				lstrcpyW(lpszNewStringW, (LPWSTR)lpszText);
			}
		}
	END_CODE_JUDGE:
		GlobalFree(lpszText);
		std::wstring strResult;
		try
		{
			strResult = std::regex_replace(lpszNewStringW, std::wregex(lpszFindText), lpszToText);
		}
		catch (...)
		{
			MessageBox(hWnd, TEXT("正規表現に問題があります。"), 0, 0);
			goto CLOSE_FILE;
		}
		SetFilePointer(hFile, 0, 0, FILE_BEGIN);
		const BYTE BOM[] = { 0xFF, 0xFE };
		DWORD dwWriteSize;
		WriteFile(hFile, BOM, 2, &dwWriteSize, 0);
		WriteFile(hFile, strResult.c_str(), (DWORD)(strResult.size() * sizeof(WCHAR)), &dwWriteSize, 0);
		SetEndOfFile(hFile);
	CLOSE_FILE:
		GlobalFree(lpszNewStringW);
		CloseHandle(hFile);
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit1;
	static HWND hEdit2;
	switch (msg)
	{
	case WM_CREATE:
		hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
		{
			DWORD dwSize1 = GetWindowTextLengthW(hEdit1);
			LPWSTR lpszText1 = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (dwSize1 + 1));
			GetWindowTextW(hEdit1, lpszText1, dwSize1 + 1);
			DWORD dwSize2 = GetWindowTextLengthW(hEdit2);
			LPWSTR lpszText2 = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (dwSize2 + 2));
			GetWindowTextW(hEdit2, lpszText2, dwSize2 + 2);
			const int nFileCount = DragQueryFile((HDROP)wParam, -1, NULL, 0);
			for (int i = 0; i<nFileCount; i++)
			{
				TCHAR szFilePath[MAX_PATH];
				DragQueryFile((HDROP)wParam, i, szFilePath, _countof(szFilePath));
				Replace(hWnd, szFilePath, lpszText1, lpszText2);
			}
			DragFinish((HDROP)wParam);
			GlobalFree(lpszText1);
			GlobalFree(lpszText2);
		}
		break;
	case WM_SIZE:
		MoveWindow(hEdit1, 10, 10, LOWORD(lParam) - 20, 32, TRUE);
		MoveWindow(hEdit2, 10, 50, LOWORD(lParam) - 20, 32, TRUE);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドロップされたテキストファイルの内容を正規表現で置換"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
