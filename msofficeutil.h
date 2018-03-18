//Copyright (c) 2018 mmYYmmdd
#pragma once
#include <fstream>
#include <string>
#include <algorithm>
#include <OleAuto.h>
#include <tchar.h>
#include <memory>
#include <type_traits>

#define SECURITY_WIN32
#include <Security.h>
#pragma comment(lib, "secur32.lib")

#if _MSC_VER < 1900
#define noexcept throw()
#endif

namespace mymd  {

    //VARIANT値の初期化
    VARIANT iVariant(VARTYPE t = VT_EMPTY) noexcept;

    //VARIANT値からの文字列取得
    BSTR getBSTR(VARIANT const&) noexcept;

    //VARIANT値からの文字列取得
    BSTR getBSTR(VARIANT const*) noexcept;

    //文字列からVT_BSTR型VARIANT値の生成
    VARIANT bstrVariant(wchar_t const* s) noexcept;

    //文字列からVT_BSTR型VARIANT値の生成
    VARIANT bstrVariant(std::wstring const& s) noexcept;

    // SafeArrayAccessData => SafeArrayUnaccessData を行うRAII
    struct SafeArrayUnaccessor {
        void operator()(SAFEARRAY* ptr) const  noexcept
        { ::SafeArrayUnaccessData(ptr); }
    };

    using safearrayRAII = std::unique_ptr<SAFEARRAY, SafeArrayUnaccessor>;

    // イテレータからのVARIANT配列生成
    template <typename InputIterator, typename F>
    VARIANT range2VArray(InputIterator begin, InputIterator end, F&& trans) noexcept
    {
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(std::distance(begin, end)), 0 };
        safearrayRAII pArray{::SafeArrayCreate(VT_VARIANT, 1, &rgb)};
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)            return iVariant();
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        std::size_t i{0};
        try
        {
            for (auto p = begin; p != end; ++p, ++i)
                std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), std::forward<F>(trans)(*p));
            auto ret = iVariant(VT_ARRAY | VT_VARIANT);
            ret.parray = pArray.get();
            return ret;
        }
        catch (const std::exception&)
        {
            return iVariant();
        }
    }

    // イテレータからの VARIANT配列生成
    template <typename InputIterator>
    VARIANT range2VArray(InputIterator begin, InputIterator end) noexcept
    {
        auto trans = [](decltype(*begin) x) -> decltype(*begin)  {
            return x;
        };
        return range2VArray(begin, end, trans);
    }

}   //namespace mymd
