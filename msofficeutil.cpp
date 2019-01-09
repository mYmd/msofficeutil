//Copyright (c) 2017 mmYYmmdd
#include "stdafx.h"
#include "msofficeutil.h"

namespace mymd  {

    //VARIANT値の初期化
    VARIANT iVariant(VARTYPE t) noexcept
    {
        VARIANT ret;
        ::VariantInit(&ret);
        ret.vt = t;
        return ret;
    }

    //VARIANT値からの文字列取得
    BSTR getBSTR(VARIANT const& expr) noexcept
    {
        if (expr.vt & VT_BYREF)
            return (VT_BSTR == (expr.vt & VT_TYPEMASK) && expr.pbstrVal) ? *expr.pbstrVal : nullptr;
        else
            return (VT_BSTR == (expr.vt & VT_TYPEMASK) && expr.bstrVal) ? expr.bstrVal : nullptr;
    }

    //VARIANT値からの文字列取得
    BSTR getBSTR(VARIANT const* expr) noexcept
    {
        return expr? getBSTR(*expr): nullptr;
    }

    //文字列からVT_BSTR型VARIANT値の生成
    VARIANT bstrVariant(wchar_t const* s) noexcept
    {
        VARIANT ret = iVariant(VT_BSTR);
        if ( s )    ret.bstrVal = ::SysAllocString(s);
        return ret;
    }

    //文字列からVT_BSTR型VARIANT値の生成
    VARIANT bstrVariant(std::wstring const& s) noexcept
    {   
        return bstrVariant(s.data());
    }
    
    //************************************************************************

    // SafeArray要素のアクセス
    safearrayRef::safearrayRef(VARIANT const& v) noexcept
        :psa{ nullptr }, pvt{ 0 }, dim{ 0 }, elemsize{ 0 }, it{ nullptr }, size({ 1, 1, 1 })
    {
        ::VariantInit(&val_);
        if (0 == (VT_ARRAY & v.vt))             return;
        psa = (0 == (VT_BYREF & v.vt)) ? v.parray : *v.pparray;
        if (!psa)                               return;
        //このAPIのせいでreinterpret_cast
        ::SafeArrayAccessData(psa, reinterpret_cast<void**>(&it));
        dim = ::SafeArrayGetDim(psa);
        if (!it || 3 < dim)
        {
            size[0] = 0;
            return;
        }
        elemsize = ::SafeArrayGetElemsize(psa);
        ::SafeArrayGetVartype(psa, &pvt);
        val_.vt = pvt | VT_BYREF;   //ここ
        for (decltype(dim) i = 0; i < dim; ++i)
        {
            LONG ub = 0, lb = 0;
            ::SafeArrayGetUBound(psa, static_cast<UINT>(i + 1), &ub);
            ::SafeArrayGetLBound(psa, static_cast<UINT>(i + 1), &lb);
            size[i] = ub - lb + 1;
        }
    }

    safearrayRef::~safearrayRef()
    {
        if (psa)     ::SafeArrayUnaccessData(psa);
        ::VariantClear(&val_);
    }

    std::size_t safearrayRef::getDim() const noexcept
    {
        return dim;
    }

    std::size_t safearrayRef::getSize(std::size_t i) const noexcept
    {
        return size[i - 1];
    }

    std::size_t safearrayRef::getOriginalLBound(std::size_t i) const noexcept
    {
        LONG lb = 0;
        ::SafeArrayGetLBound(psa, static_cast<UINT>(i), &lb);
        return lb;
    }

    VARIANT& safearrayRef::operator()(std::size_t i, std::size_t j, std::size_t k) noexcept
    {
        auto distance = size[0] * size[1] * k + size[0] * j + i;
        if (pvt == VT_VARIANT)
        {
            return *reinterpret_cast<VARIANT*>(it + distance * elemsize);
        }
        else
        {
            val_.pvarVal = reinterpret_cast<VARIANT*>(it + distance * elemsize);
            return val_;
        }
    }

}   //namespace mymd
