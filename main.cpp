#include <windows.h>
#include <shlobj.h>
#include <gdiplus.h>
#include <iostream>
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")

using namespace Gdiplus;

// Custom message to signal Ctrl+V
#define WM_PASTE_SCREENSHOT (WM_USER + 1)

// Get CLSID for image encoder
int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
    UINT num = 0, size = 0;
    GetImageEncodersSize(&num, &size);
    if (size == 0) return -1;

    ImageCodecInfo* pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    GetImageEncoders(num, size, pImageCodecInfo);

    for (UINT j = 0; j < num; ++j) {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;
        }
    }
    free(pImageCodecInfo);
    return -1;
}

// Save clipboard bitmap to file
bool SaveClipboardImageToFile(const wchar_t* directory) {
    std::wcout << L"Attempting to save clipboard image...\n";
    if (!OpenClipboard(NULL)) {
        std::wcout << L"Failed to open clipboard!\n";
        return false;
    }

    HBITMAP hBitmap = (HBITMAP)GetClipboardData(CF_BITMAP);
    if (!hBitmap) {
        std::wcout << L"No bitmap found in clipboard!\n";
        CloseClipboard();
        return false;
    }

    GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    Bitmap* bitmap = Bitmap::FromHBITMAP(hBitmap, NULL);
    if (!bitmap) {
        std::wcout << L"Failed to create bitmap from HBITMAP!\n";
        CloseClipboard();
        return false;
    }

    wchar_t filepath[MAX_PATH];
    SYSTEMTIME st;
    GetLocalTime(&st);
    swprintf_s(filepath, MAX_PATH, L"%s\\Screenshot_%04d%02d%02d_%02d%02d%02d.png",
               directory, st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

    CLSID pngClsid;
    GetEncoderClsid(L"image/png", &pngClsid);
    Status stat = bitmap->Save(filepath, &pngClsid, NULL);

    if (stat == Ok) {
        std::wcout << L"Saved image to: " << filepath << L"\n";
    } else {
        std::wcout << L"Failed to save image! Status: " << stat << L"\n";
    }

    delete bitmap;
    GdiplusShutdown(gdiplusToken);
    CloseClipboard();
    return stat == Ok;
}

// Get current File Explorer directory
bool GetExplorerDirectory(wchar_t* outPath, size_t size) {
    std::wcout << L"Checking for active Explorer window...\n";
    HWND hwnd = GetForegroundWindow();
    if (!hwnd) {
        std::wcout << L"No foreground window found!\n";
        return false;
    }

    IShellWindows* pShellWindows;
    HRESULT hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&pShellWindows);
    if (FAILED(hr)) {
        std::wcout << L"Failed to create IShellWindows! HR: " << hr << L"\n";
        return false;
    }

    long count;
    pShellWindows->get_Count(&count);
    std::wcout << L"Found " << count << L" shell windows.\n";

    for (long i = 0; i < count; i++) {
        VARIANT v;
        V_VT(&v) = VT_I4;
        V_I4(&v) = i;
        IDispatch* pDisp;
        hr = pShellWindows->Item(v, &pDisp);
        if (FAILED(hr) || !pDisp) continue;

        IWebBrowserApp* pWebBrowserApp;
        hr = pDisp->QueryInterface(IID_IWebBrowserApp, (void**)&pWebBrowserApp);
        pDisp->Release();
        if (FAILED(hr)) continue;

        HWND hwndBrowser;
        pWebBrowserApp->get_HWND((LONG_PTR*)&hwndBrowser);
        if (hwndBrowser == hwnd) {
            IServiceProvider* pServiceProvider;
            hr = pWebBrowserApp->QueryInterface(IID_IServiceProvider, (void**)&pServiceProvider);
            pWebBrowserApp->Release();
            if (FAILED(hr)) continue;

            IShellBrowser* pShellBrowser;
            hr = pServiceProvider->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&pShellBrowser);
            pServiceProvider->Release();
            if (FAILED(hr)) continue;

            IShellView* pShellView;
            hr = pShellBrowser->QueryActiveShellView(&pShellView);
            pShellBrowser->Release();
            if (FAILED(hr)) continue;

            IFolderView* pFolderView;
            hr = pShellView->QueryInterface(IID_IFolderView, (void**)&pFolderView);
            pShellView->Release();
            if (FAILED(hr)) continue;

            IPersistFolder2* pPersistFolder;
            hr = pFolderView->GetFolder(IID_IPersistFolder2, (void**)&pPersistFolder);
            pFolderView->Release();
            if (FAILED(hr)) continue;

            LPITEMIDLIST pidl;
            hr = pPersistFolder->GetCurFolder(&pidl);
            pPersistFolder->Release();
            if (SUCCEEDED(hr)) {
                SHGetPathFromIDListW(pidl, outPath);
                CoTaskMemFree(pidl);
                std::wcout << L"Found Explorer directory: " << outPath << L"\n";
                return true;
            }
        }
        pWebBrowserApp->Release();
    }
    pShellWindows->Release();
    std::wcout << L"No matching Explorer window found!\n";
    return false;
}

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION && wParam == WM_KEYDOWN) {
        PKBDLLHOOKSTRUCT p = (PKBDLLHOOKSTRUCT)lParam;
        if (p->vkCode == 'V' && (GetAsyncKeyState(VK_CONTROL) & 0x8000)) {
            std::wcout << L"Ctrl+V detected! Sending message to main thread...\n";
            // Post a message to the main thread instead of doing work here
            PostThreadMessage(GetCurrentThreadId(), WM_PASTE_SCREENSHOT, 0, 0);
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

int main() {
    std::wcout << L"Starting program...\n";
    CoInitialize(NULL);

    HHOOK hook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandle(NULL), 0);
    if (!hook) {
        std::wcout << L"Failed to install keyboard hook! Error: " << GetLastError() << L"\n";
        CoUninitialize();
        return 1;
    }
    std::wcout << L"Keyboard hook installed. Press Ctrl+V in Explorer to save a screenshot.\n";

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_PASTE_SCREENSHOT) {
            std::wcout << L"Received paste screenshot message!\n";
            wchar_t directory[MAX_PATH];
            if (GetExplorerDirectory(directory, MAX_PATH)) {
                if (SaveClipboardImageToFile(directory)) {
                    std::wcout << L"Success!\n";
                } else {
                    std::wcout << L"Failed to save image!\n";
                }
            } else {
                std::wcout << L"Failed to get Explorer directory!\n";
            }
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnhookWindowsHookEx(hook);
    CoUninitialize();
    std::wcout << L"Program exiting...\n";
    return 0;
}
