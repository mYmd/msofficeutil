//Copyright (c) 2018 mmYYmmdd
#pragma once
#include "msofficeutil.h"
#include <regex>
#include <vector>

namespace mymd  {

    //肯定後読み、否定後読みを実装した正規表現
    //ただし (?<= と (?<! はパターンの先頭に限る
    class regexp2   {
        int positive_negative;
        std::wstring p_ahead, p_behind;
        std::wstring match_, prefix_, suffix_;
        std::vector<std::wstring> submatches_;
    public:
        regexp2(std::wstring const& pattern);
        void execute(std::wstring const& target, bool icase, std::wstring const* original = nullptr);
        std::wstring const& match() const;
        std::wstring const& prefix() const;
        std::wstring const& suffix() const;
        std::vector<std::wstring> const& submatches() const;
    };

}