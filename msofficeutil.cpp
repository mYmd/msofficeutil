//Copyright (c) 2017 mmYYmmdd
#include "stdafx.h"
#include "msofficeutil.h"

namespace mymd  {

    //VARIANT�l�̏�����
    VARIANT iVariant(VARTYPE t) noexcept
    {
        VARIANT ret;
        ::VariantInit(&ret);
        ret.vt = t;
        return ret;
    }

    //VARIANT�l����̕�����擾
    BSTR getBSTR(VARIANT const& expr) noexcept
    {
        if (expr.vt & VT_BYREF)
            return ((expr.vt & VT_BSTR) && expr.pbstrVal) ? *expr.pbstrVal : nullptr;
        else
            return ((expr.vt & VT_BSTR) && expr.bstrVal) ? expr.bstrVal : nullptr;
    }

    //VARIANT�l����̕�����擾
    BSTR getBSTR(VARIANT const* expr) noexcept
    {
        return expr? getBSTR(*expr): nullptr;
    }

    //�����񂩂�VT_BSTR�^VARIANT�l�̐���
    VARIANT bstrVariant(wchar_t const* s) noexcept
    {
        VARIANT ret = iVariant(VT_BSTR);
        if ( s )    ret.bstrVal = ::SysAllocString(s);
        return ret;
    }

    //�����񂩂�VT_BSTR�^VARIANT�l�̐���
    VARIANT bstrVariant(std::wstring const& s) noexcept
    {   
        return bstrVariant(s.data());
    }
    
}   //namespace mymd
