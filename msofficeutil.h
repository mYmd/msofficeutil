//Copyright (c) 2018 mmYYmmdd
#pragma once
#include <fstream>
#include <array>
#include <string>
#include <algorithm>
#include <OleAuto.h>
#include <tchar.h>
#include <memory>

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

    struct swap_v_t {
        void operator ()(VARIANT& a, VARIANT& b) const noexcept     { std::swap(a, b); }
        void operator ()(VARIANT& a, VARIANT&& b) const noexcept    { std::swap(a, b); }
    };

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
        swap_v_t swap_v;
        try
        {
            for (auto p = begin; p != end; ++p, ++i)
                swap_v(*reinterpret_cast<VARIANT*>(it + i * elemsize), std::forward<F>(trans)(*p));
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

    //********************************************************

    // SafeArray要素のアクセス
    class safearrayRef {
        SAFEARRAY*      psa;
        VARTYPE         pvt;
        std::size_t     dim;
        std::size_t     elemsize;
        char*           it;
        VARIANT         val_;
        std::array<std::size_t, 3>  size;
    public:
        explicit safearrayRef(VARIANT const& v) noexcept;
        ~safearrayRef();
        safearrayRef(safearrayRef const&) = delete;
        safearrayRef(safearrayRef&&) = delete;
        std::size_t getDim() const noexcept;
        std::size_t getSize(std::size_t i) const noexcept;
        std::size_t getOriginalLBound(std::size_t i) const noexcept;
        VARIANT& operator()(std::size_t i, std::size_t j = 0, std::size_t k = 0) noexcept;
    };

}   //namespace mymd
