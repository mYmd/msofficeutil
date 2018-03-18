//Copyright (c) 2018 mmYYmmdd
#include "stdafx.h"
#include "msofficeutil.h"
#include "regExp2.hpp"
#include <array>

//肯定後読み、否定後読みを実装した正規表現
//ただし (?<= と (?<! はパターンの先頭に限る

//配列 {match, prefix, suffix, submatch()} を返す
VARIANT __stdcall
RegExpExec(VARIANT const& ptrn, VARIANT const& target, __int8 icase)
{
    auto const pattern = mymd::getBSTR(ptrn);
    auto const targetStr = mymd::getBSTR(target);
    if ( !pattern || !targetStr )   return mymd::iVariant();
    mymd::regexp2 re2{pattern};
    re2.execute(targetStr, icase != 0);
    auto const& submatches = re2.submatches();
    std::vector<VARIANT> sub_vec;
    auto n = submatches.size();
    sub_vec.reserve(n);
    for ( std::size_t i = 0; i < n; ++i )
    {
        sub_vec.push_back(mymd::bstrVariant(submatches[i]));
    }
    std::array<VARIANT, 4> arr = { mymd::bstrVariant(re2.match())   ,
                                   mymd::bstrVariant(re2.prefix())  ,
                                   mymd::bstrVariant(re2.suffix())  ,
                                   mymd::range2VArray(sub_vec.begin(), sub_vec.end()) };
    return mymd::range2VArray(arr.begin(), arr.end());
}

namespace mymd  {

    regexp2::regexp2(std::wstring const& pattern) : positive_negative{0}
    {
        if ( pattern.size() < 6 )
        {
            p_behind = pattern;
            return;
        }
        positive_negative = (pattern.compare(0, 4, L"(?<=") == 0)? 1:
            ((pattern.compare(0, 4, L"(?<!") == 0)? -1: 0);
        if ( 0 == positive_negative )
        {
            p_behind = pattern;
        }
        else
        {
            int count = 1;
            auto it = pattern.begin() + 4;
            for ( ; it < pattern.end() && 0 < count; ++it )
            {
                if ( *it == L'(' )  ++count;
                if ( *it == L')' )  --count;
            }
            if ( 0 < count )
            {
                positive_negative = 0;
                p_behind = pattern;
            }
            else
            {
                p_ahead = pattern.substr(4, it-pattern.begin()-4-1) + L'$';
                p_behind = pattern.substr(it-pattern.begin(), pattern.end()-it);
            }
        }
    }

    std::wstring const& regexp2::match() const
    {
        return match_;
    }

    std::wstring const& regexp2::prefix() const
    {
        return prefix_;
    }

    std::wstring const& regexp2::suffix() const
    {
        return suffix_;
    }

    std::vector<std::wstring> const& regexp2::submatches() const
    {
        return submatches_;
    }

    void regexp2::execute(std::wstring const& target, bool icase)
    {
        prefix_.clear();
        suffix_.clear();
        match_.clear();
        submatches_.clear();
        try    {
            std::wsmatch wm_behind;
            auto type = std::regex_constants::ECMAScript;
            if ( icase )    type |= std::regex_constants::icase;
            std::wregex wrg2{p_behind, type};
            if ( 0 == positive_negative )
            {
                std::regex_search(target, wm_behind, wrg2);
                prefix_ = std::wstring{wm_behind.prefix()};
            }
            else
            {
                bool const positive = (positive_negative == 1);
                std::wregex wrg1{p_ahead, type};
                std::wsmatch wm_ahead;
                bool match_behind = std::regex_search(target, wm_behind, wrg2);
                prefix_ = wm_behind.prefix();
                bool match_ahead;
                while ( match_behind && (wm_behind[0].first < target.end()) )
                {
                    match_ahead = std::regex_search(prefix_, wm_ahead, wrg1);
                    if ( match_ahead == positive )   break;
                    auto first = wm_behind[0].first + 1;
                    match_behind = std::regex_search(first, target.end(), wm_behind, wrg2);
                    prefix_ = std::wstring{target.begin(), wm_behind.prefix().second};
                }
                if ( match_ahead != positive )       wm_behind.swap(std::wsmatch{});    //初期化
            }
            match_ = std::wstring{wm_behind.str(0)};
            suffix_ = std::wstring{wm_behind.suffix()};
            submatches_.clear();
            auto n = wm_behind.size();
            for ( std::size_t i = 1; i < n; ++i )
            {
                submatches_.push_back(wm_behind.str(i));
            }
        }
        catch (const std::exception&)   {
            return ;
        }
    }

}
