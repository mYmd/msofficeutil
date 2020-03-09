//csvmap_vba.cpp
//Copyright (c) 2019 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include <vector>
#include "csvmap.hpp"

//CSVファイルの読み込み
// stream : 入力ストリーム（fstream または stringstream を想定 ）
// R : 出力する文字型（wchar_tを想定）
// S : 対象streamの文字型（ファイルをcodeconvertする前提ならstringを想定）
// delimiter : 区切り文字（これによって型R が決定される）
// elem_func   : 要素毎のコールバック　　[&](std::size_t count, std::wstring&& expr) 等、 (項目番号, 要素文字列)
// record_func : 1レコード読終時のコールバック　[&](std::size_t count, std::size_t size)　等、(行番号, 当該行の項目数)
// codepage : CP_UTF8, CP_ACP など
// 返り値 : 読んだ行数（テキストファイルの場合は最後の空行もカウント）

VARIANT __stdcall Map_CSV(VARIANT const&    Path        ,
                          VARIANT const&    Delimiter   ,
                          __int32 const     Code_Page   ,
                          __int32 const     Head_N      ,
                          __int32 const     Head_Cut    )
{
    auto path = mymd::getBSTR(Path);
    auto delimiter= mymd::getBSTR(Delimiter);
    if (!path || !delimiter)    return mymd::iVariant();
    std::ifstream ifs(path);
    if (ifs.bad())              return mymd::iVariant();
    std::vector<std::wstring> record;
    std::vector<VARIANT> vec;
    auto elem_func = [&](std::size_t count, std::wstring&& expr)    {
        record.push_back(std::move(expr));
    };
    auto head_cut = static_cast<std::size_t>(Head_Cut);
    auto head_n = static_cast<std::size_t>(Head_N < 0? INT_MAX: Head_N);
    auto record_func = [&](std::size_t count, std::size_t size) {
        if (head_cut <= count && count < head_n)
        {
            auto elem = mymd::range2VArray(record.begin(), record.end(), [](std::wstring& w) { return mymd::bstrVariant(w); });
            vec.push_back(elem);
            record.clear();
            return true;
        }
        return false;
    };
    auto s = mymd::map_csv(ifs, delimiter[0], elem_func, record_func, Code_Page);
    return mymd::range2VArray(vec.begin(), vec.end());
}
