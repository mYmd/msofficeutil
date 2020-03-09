//FileSystem.cpp
//Copyright (c) 2019 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include <fileapi.h>
#include <ctime>
#include "Objbase.h"

namespace   {

// "yyyy/mm/dd hh:mm:ss" 形式の文字列(local time)からFILETIME構造体(UTC)を返す
bool getFileTime_t(wchar_t const* const expr, FILETIME& outVar)
{
    if (!expr || std::wcslen(expr) < 10)        return false;
    ::SYSTEMTIME sTime{
        static_cast<WORD>(std::wcstol(expr, nullptr, 10)),      //wYear
        static_cast<WORD>(std::wcstol(expr + 5, nullptr, 10)),  //wMonth
        0,                                                      //wDayOfWeek (ignored)
        static_cast<WORD>(std::wcstol(expr + 8, nullptr, 10)),  //wDay
        0,                              //wHour
        0,                              //wMinute
        0,                              //wSecond
        0                               //wMilliseconds
    };
    if (18 < std::wcslen(expr))
    {
        sTime.wHour   = static_cast<WORD>(std::wcstol(expr + 11, nullptr, 10));
        sTime.wMinute = static_cast<WORD>(std::wcstol(expr + 14, nullptr, 10));
        sTime.wSecond = static_cast<WORD>(std::wcstol(expr + 17, nullptr, 10));
    }
    FILETIME localFileTime;
    return TRUE == ::SystemTimeToFileTime(&sTime, &localFileTime)
        && TRUE == ::LocalFileTimeToFileTime(&localFileTime, &outVar);
}

// FILETIME構造体(UTC)から "yyyy/mm/dd hh:mm:ss" 形式(local time)の文字列を表すVARIANTを返す
VARIANT getFileTimeVariant(FILETIME const& ftime)
{
    FILETIME  ftime_local;
    if (FALSE == ::FileTimeToLocalFileTime(&ftime, &ftime_local))   return mymd::iVariant();
    ::SYSTEMTIME sTime;
    if (FALSE == ::FileTimeToSystemTime(&ftime_local, &sTime)) return mymd::iVariant();
    wchar_t buf[] = L"yyyy/mm/dd hh:mm:ss";
    ::swprintf_s(buf, L"%04d/%02d/%02d %02d:%02d:%02d", sTime.wYear, sTime.wMonth, sTime.wDay, sTime.wHour, sTime.wMinute, sTime.wSecond);
    return mymd::bstrVariant(buf);
}

}

/*
Time Functions  https://docs.microsoft.com/ja-jp/windows/desktop/SysInfo/time-functions
    https://docs.microsoft.com/ja-jp/windows/desktop/api/minwinbase/ns-minwinbase-filetime
Changing a File Time to the Current Time
    https://docs.microsoft.com/ja-jp/windows/desktop/SysInfo/changing-a-file-time-to-the-current-time
*/

//ファイルの作成日時、最終更新日時、最終アクセス日時を変更する
//  "yyyy/mm/dd hh:mm:ss" 形式のString
__int32 __stdcall
changeFileTimeStamp(VARIANT const& FilePath, VARIANT const& CreationTime, VARIANT const& LastWriteTime, VARIANT const& LastAccessTime)
{
    auto filePath = mymd::getBSTR(FilePath);
    FILETIME creationTime, lastAccessTime, lastWriteTime;
    if (!filePath   ||
        !getFileTime_t(mymd::getBSTR(CreationTime), creationTime)       ||
        !getFileTime_t(mymd::getBSTR(LastWriteTime), lastWriteTime)     ||
        !getFileTime_t(mymd::getBSTR(LastAccessTime), lastAccessTime)   )        return 0;
    auto fileHandle = ::CreateFileW(filePath,
                                    FILE_WRITE_ATTRIBUTES,
                                    FILE_SHARE_WRITE,
                                    nullptr,
                                    OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL,
                                    nullptr);
    return (fileHandle == INVALID_HANDLE_VALUE)? 0
        : (TRUE == ::SetFileTime(fileHandle, &creationTime, &lastAccessTime,  &lastWriteTime))? 1: 0;
}

//ファイルの作成日時、最終更新日時、最終アクセス日時を取得する
//  "yyyy/mm/dd hh:mm:ss"形式のString
VARIANT __stdcall
getFileTimeStamp(VARIANT const& FilePath)
{
    auto filePath = mymd::getBSTR(FilePath);
    if (!filePath)        return mymd::iVariant();
    auto fileHandle = ::CreateFileW(filePath,
                                    FILE_READ_ATTRIBUTES,
                                    FILE_SHARE_READ,
                                    nullptr,
                                    OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL,
                                    nullptr);
    if (fileHandle != INVALID_HANDLE_VALUE)
    {
        FILETIME creationTime, lastAccessTime, lastWriteTime;
        if (TRUE == ::GetFileTime(fileHandle, &creationTime, &lastAccessTime,  &lastWriteTime))
        {
            VARIANT ret[3] = { getFileTimeVariant(creationTime), getFileTimeVariant(lastWriteTime), getFileTimeVariant(lastAccessTime)};
            return mymd::range2VArray(std::begin(ret), std::end(ret));
        }
        else    return mymd::iVariant();
    }
    else    return mymd::iVariant();
}

