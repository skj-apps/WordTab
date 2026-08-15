// WordTab - file logging.
//
// Deliberately dumb: one file, appended, no history. It answers "what did the last run do",
// which is the only question worth asking from inside a host process we do not control.

#include "wordtab.h"
#include <shlobj.h>
#include <stdio.h>
#include <stdarg.h>

static const LONGLONG kMaxBytes = 512 * 1024;

static CRITICAL_SECTION g_logLock;
static BOOL g_logLockReady = FALSE;
static wchar_t g_logPath[MAX_PATH] = L"";
static wchar_t g_logDir[MAX_PATH] = L"";

static void EnsurePaths(void)
{
    if (g_logPath[0] != L'\0')
        return;

    wchar_t localAppData[MAX_PATH];
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, localAppData)))
        return;

    _snwprintf(g_logDir, MAX_PATH, L"%s\\WordTab", localAppData);
    g_logDir[MAX_PATH - 1] = L'\0';
    CreateDirectoryW(g_logDir, NULL);

    _snwprintf(g_logPath, MAX_PATH, L"%s\\wordtab.log", g_logDir);
    g_logPath[MAX_PATH - 1] = L'\0';
}

const wchar_t* LogFilePath(void)
{
    EnsurePaths();
    return g_logPath;
}

// Roll rather than grow without bound.
static void RollIfLarge(void)
{
    WIN32_FILE_ATTRIBUTE_DATA info;
    if (!GetFileAttributesExW(g_logPath, GetFileExInfoStandard, &info))
        return;

    LARGE_INTEGER size;
    size.HighPart = (LONG)info.nFileSizeHigh;
    size.LowPart = info.nFileSizeLow;
    if (size.QuadPart > kMaxBytes)
        DeleteFileW(g_logPath);
}

void LogWrite(const wchar_t* format, ...)
{
    if (!g_logLockReady)
    {
        // Racy only in theory: the first LogWrite happens on Word's single loading thread, long
        // before we start any thread of our own.
        InitializeCriticalSection(&g_logLock);
        g_logLockReady = TRUE;
    }

    EnterCriticalSection(&g_logLock);

    EnsurePaths();
    if (g_logPath[0] == L'\0')
    {
        LeaveCriticalSection(&g_logLock);
        return;
    }

    RollIfLarge();

    wchar_t message[2048];
    va_list args;
    va_start(args, format);
    _vsnwprintf(message, 2048, format, args);
    va_end(args);
    message[2047] = L'\0';

    SYSTEMTIME now;
    GetLocalTime(&now);

    wchar_t line[2400];
    _snwprintf(line, 2400,
               L"%04d-%02d-%02d %02d:%02d:%02d.%03d  pid=%-6lu tid=%-6lu  %s\r\n",
               now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, now.wSecond,
               now.wMilliseconds, GetCurrentProcessId(), GetCurrentThreadId(), message);
    line[2399] = L'\0';

    // UTF-8 on disk so the file opens cleanly in any editor.
    int bytes = WideCharToMultiByte(CP_UTF8, 0, line, -1, NULL, 0, NULL, NULL);
    if (bytes > 1)
    {
        char* utf8 = (char*)HeapAlloc(GetProcessHeap(), 0, (SIZE_T)bytes);
        if (utf8)
        {
            WideCharToMultiByte(CP_UTF8, 0, line, -1, utf8, bytes, NULL, NULL);

            HANDLE file = CreateFileW(g_logPath, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE,
                                      NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            if (file != INVALID_HANDLE_VALUE)
            {
                DWORD written = 0;
                // bytes - 1 drops the terminating NUL.
                WriteFile(file, utf8, (DWORD)(bytes - 1), &written, NULL);
                CloseHandle(file);
            }
            HeapFree(GetProcessHeap(), 0, utf8);
        }
    }

    LeaveCriticalSection(&g_logLock);
}
