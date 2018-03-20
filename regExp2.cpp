//Copyright (c) 2018 mmYYmmdd
#include "stdafx.h"
#include "msofficeutil.h"
#include "regExp2.hpp"
#include <array>
#include <algorithm>

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
    re2.execute(targetStr, icase != 0, nullptr);
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

    std::locale stdLocale = std::locale::classic();

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

    //
    std::wstring getOriginal(std::wstring const&                 target  , 
                             std::wstring const*                 original,
                             std::wsmatch::value_type::iterator  begin   ,
                             std::wsmatch::value_type::iterator  end     )
    {
        return ( begin < end )? original->substr(begin - target.begin(), end - begin): std::wstring{};
    }

    void regexp2::execute(std::wstring const& target, bool icase, std::wstring const* original)
    {
        // VC++のregexのバグをworkaround
        // icase 指定時に[a-z]$ も [A-Z]$ も "ABC" にマッチしない問題
        //  ==> icase 指定時は target を小文字に揃えてしまう
        if ( icase )
        {
            auto lowerTarget(target);
            std::transform(lowerTarget.begin(), lowerTarget.end(), lowerTarget.begin(),
                           [](wchar_t x) { return std::tolower(x, stdLocale);} );
            execute(lowerTarget, false, &target);   //再帰
            return;
        }
        prefix_.clear();
        suffix_.clear();
        match_.clear();
        submatches_.clear();
        try    {
            using iterator = std::wsmatch::value_type::iterator;
            iterator prefix_begin{};
            iterator prefix_end{};
            std::wsmatch wm_behind;
            auto type = std::regex_constants::ECMAScript;
            if ( original )     type |= std::regex_constants::icase;    //再帰呼び出しの場合
            std::wregex wrg_behind{p_behind, type};
            bool match_total = false;
            if ( 0 == positive_negative )
            {
                match_total = std::regex_search(target, wm_behind, wrg_behind);
                prefix_begin = wm_behind.prefix().first;
                prefix_end = wm_behind.prefix().second;
            }
            else
            {
                bool const positive = (positive_negative == 1);
                std::wregex wrg_ahead{p_ahead, type};
                std::wsmatch wm_ahead;
                bool match_behind = std::regex_search(target, wm_behind, wrg_behind);
                prefix_begin = wm_behind.prefix().first;
                prefix_end = wm_behind.prefix().second;
                bool match_ahead;
                while ( match_behind && (wm_behind[0].first < target.end()) )
                {
                    match_ahead = std::regex_search(prefix_begin, prefix_end, wm_ahead, wrg_ahead);
                    if ( match_ahead == positive )
                    {
                        match_total = true;
                        break;
                    }
                    auto first = wm_behind[0].first + 1;
                    match_behind = std::regex_search(first, target.end(), wm_behind, wrg_behind);
                    prefix_begin = target.begin();
                    prefix_end = wm_behind[0].first;
                }
            }
            if ( !original )    original = &target;
            if ( match_total )
            {
                match_ = getOriginal(target, original, wm_behind[0].first, wm_behind[0].second);
                prefix_ = getOriginal(target, original, prefix_begin, prefix_end);
                suffix_ = getOriginal(target, original, wm_behind.suffix().first, wm_behind.suffix().second);
                auto n = wm_behind.size();
                submatches_.reserve(n-1);
                for ( std::size_t i = 1; i < n; ++i )
                {
                    submatches_.push_back(getOriginal(target, original, wm_behind[i].first, wm_behind[i].second));
                }
            }
        }
        catch (const std::exception&)   {
            return;
        }
    }

}
