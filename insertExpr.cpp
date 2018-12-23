//insertExpr.cpp
//Copyright (c) 2018 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include <vector>

namespace mymd {

    template <typename V>
    std::wstring make_insert_expr_imple(std::wstring, safearrayRef& attr, V&&);

    // MERGE INTO **** A USING(SELECT ** AS **, ** AS **, ...) B
    template <typename V>
    std::wstring    make_using_part(wchar_t const*      schema_table_name,
                                    std::wstring const& A_name  ,
                                    std::wstring const& B_name  ,
                                    safearrayRef&       attr    ,
                                    V&&                 values  ,
                                    bool                for_oracle);

    // ON (A.***=B.*** AND A.***=B.***)
    std::wstring    make_ON_expr(std::wstring const&    A_name,
                                 std::wstring const&    B_name,
                                 safearrayRef&          attr,
                                 safearrayRef&          key);

    //  WHEN MATCHED THEN UPDATE SET *** = B.***, ...
    std::wstring    make_set_part(std::wstring const& B_name, safearrayRef& attr, safearrayRef& key);

    // WHEN NOT MATCHED THEN INSERT (***,***,...)  VALUES (B.***, B.***, ...);
    std::wstring    make_values_part(std::wstring const& B_name, safearrayRef& attr);

    std::pair<std::wstring, std::wstring>
    getTemporaryTableNames(std::wstring const& name);

}       //namespace mymd

// INSERT文の生成
VARIANT __stdcall 
make_insert_expr(VARIANT const& schema_table_name_, VARIANT const& attr_, VARIANT const& values_)
{
    auto schema_table_name = mymd::getBSTR(schema_table_name_);
    if (!schema_table_name)         return mymd::iVariant();
    mymd::safearrayRef attr{ attr_ };
    mymd::safearrayRef values{ values_ };
    if ( attr.getDim()==0 || values.getDim()==0 )       return mymd::iVariant();
    std::wstring prefix{ L"INSERT INTO " };
    (prefix += schema_table_name) += L" (";
    std::vector<std::wstring>   buf;
    auto dummy = mymd::iVariant();
    switch (values.getDim())
    {
    case 1:
    {
        if (0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt))    //単一レコード
        {
            auto const bound = values.getSize(1);
            auto value_func = [&](std::size_t j) {
                return (j < bound) ? values(j) : dummy;
            };
            return mymd::bstrVariant(mymd::make_insert_expr_imple(prefix, attr, value_func));
        }
        else            //複数レコードのジャグ配列
        {
            for (std::size_t i = 0; i < values.getSize(1); ++i)
            {
                mymd::safearrayRef value2{ values(i) };
                auto const bound = value2.getSize(1);
                auto value_func = [&](std::size_t j) {
                    return (j < bound) ? value2(j) : dummy;
                };
                buf.push_back(mymd::make_insert_expr_imple(prefix, attr, value_func));
            }
        }
        break;
    }
    case 2:            //複数レコードの2次元配列
    {
        auto const bound = values.getSize(2);
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            auto value_func = [&](std::size_t j) {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(mymd::make_insert_expr_imple(prefix, attr, value_func));
        }
        break;
    }
    default:
        return mymd::iVariant();
    }
    return mymd::range2VArray(buf.begin(),
                              buf.end(), 
                              [](std::wstring const& w) { return mymd::bstrVariant(w); });
}


VARIANT __stdcall
make_merge_expr(VARIANT const&  schema_table_name_  ,
                VARIANT const&  attr_   ,
                VARIANT const&  key_    ,
                VARIANT const&  values_ ,
                __int32         oracle  )
{
    auto schema_table_name = mymd::getBSTR(schema_table_name_);
    if (!schema_table_name)         return mymd::iVariant();
    mymd::safearrayRef attr{ attr_ };
    mymd::safearrayRef key{ key_ };
    mymd::safearrayRef values{ values_ };
    if (attr.getDim() == 0 || key.getDim() == 0 || values.getDim() == 0)       return mymd::iVariant();
    std::vector<std::wstring>   buf;
    auto const altname = mymd::getTemporaryTableNames(schema_table_name);
    std::wstring const& A_name = altname.first;
    std::wstring const& B_name = altname.second;
    auto dummy = mymd::iVariant();
    auto for_oracle = (0 != oracle);
    switch (values.getDim())
    {
    case 1:
    {
        if (0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt))    //単一レコード
        {
            auto const bound = values.getSize(1);
            auto value_func = [&](std::size_t j) {
                return (j < bound) ? values(j) : dummy;
            };
            return mymd::bstrVariant( mymd::make_using_part(schema_table_name, A_name, B_name, attr, value_func, for_oracle)
                                    + mymd::make_ON_expr(A_name, B_name, attr, key)
                                    + mymd::make_set_part(B_name, attr, key)
                                    + mymd::make_values_part(B_name, attr)
                                    );
        }
        else            //複数レコードのジャグ配列
        {
            for (std::size_t i = 0; i < values.getSize(1); ++i)
            {
                mymd::safearrayRef value2{ values(i) };
                auto const bound = value2.getSize(1);
                auto value_func = [&](std::size_t j) {
                    return (j < bound) ? value2(j) : dummy;
                };
                buf.push_back(  mymd::make_using_part(schema_table_name, A_name, B_name, attr, value_func, for_oracle)
                              + mymd::make_ON_expr(A_name, B_name, attr, key)
                              + mymd::make_set_part(B_name, attr, key)
                              + mymd::make_values_part(B_name, attr)
                );
            }
        }
        break;
    }
    case 2:            //複数レコードの2次元配列
    {
        auto const bound = values.getSize(2);
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            auto value_func = [&](std::size_t j) {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(  mymd::make_using_part(schema_table_name, A_name, B_name, attr, value_func, for_oracle)
                          + mymd::make_ON_expr(A_name, B_name, attr, key)
                          + mymd::make_set_part(B_name, attr, key)
                          + mymd::make_values_part(B_name, attr)
            );
        }
        break;
    }
    default:
        return mymd::iVariant();
    }
    return mymd::range2VArray(buf.begin(),
                              buf.end(),
                              [](std::wstring const& w) { return mymd::bstrVariant(w); });
}

namespace mymd {

//浮動小数をあらわす文字列から末尾の 0（と.） を除去する　（指数表示ではない前提）
std::wstring&   trimfloat(std::wstring& fw)
{
    if (std::wstring::npos != fw.find(L'.'))    //小数点があれば
    {
        while (fw.back() == L'0')       fw.pop_back();
        if (fw.back() == L'.')          fw.pop_back();
        if (fw.empty() || fw == L"-" || fw == L"+")   fw = L'0';
    }
    return fw;
}

std::wstring&   quote_double(std::wstring& w, BSTR s)
{
    w = s;
    std::wstring::size_type pos{ std::wstring::npos };
    while (std::wstring::npos != (pos = w.find_last_of(L'\'', pos)))    // 'があれば
    {
        w.insert(pos, 1, L'\'');
        if (0 == pos--) break;
    }
    return w;
}

    class   valueof_variant {
        std::wstring value_;
    public:
        std::wstring const& value() {   return value_;  }
        bool operator ()(VARIANT& v, bool evaluate_null = false)
        {
            value_.clear();
            auto b = getBSTR(v);
            if (b)
            {
                value_ = (L'\'' + quote_double(value_, b) + L'\'');
                return true;
            }
            else
            {
                auto varDest = iVariant();
                switch (v.vt)
                {
                case VT_EMPTY: case VT_NULL:
                    if (evaluate_null)      value_ = L"NULL";
                    return  evaluate_null;
                case VT_R4:
                    value_ = trimfloat(std::to_wstring(v.fltVal));
                    return true;
                case VT_R8:
                    value_ = trimfloat(std::to_wstring(v.dblVal));
                    return true;
                case VT_I1: case VT_I2: case VT_I4: case VT_I8:
                    if (S_OK == ::VariantChangeType(&varDest, &v, 0, VT_I8))
                    {
                        value_ = std::to_wstring(varDest.llVal);
                        return true;
                    }
                    else    return false;
                case VT_DATE: case VT_CY:
                    if (S_OK == ::VariantChangeType(&varDest, &v, 0, VT_BSTR))
                    {
                        value_ = L'\'';    value_ += varDest.bstrVal;    value_ += L'\'';
                        return true;
                    }
                    else    return false;
                default:    return false;
                }
            }
        }
    };

template <typename V>
std::wstring    make_insert_expr_imple(std::wstring     prefix  ,
                                       safearrayRef&    attr    ,
                                       V&&              values  )
{
    //VT_EMPTY        VT_NULL
    std::wstring value_part{ L") VALUES (" };
    bool valid{ false };
    BSTR b;
    valueof_variant vv;
    for (std::size_t j = 0; j < attr.getSize(1); ++j)
    {
        if ( vv(std::forward<V>(values)(j)) && (b = getBSTR(attr(j))) )
        {
            prefix += b;    prefix += L',';
            value_part += vv.value() + L',';
            valid = true;
        }
    }
    if (valid)
    {
        prefix.pop_back();
        value_part.pop_back();
        return prefix + value_part + L");";
    }
    else
    {
        return std::wstring{};
    }
}

//***********************************************************************
// MERGE INTO **** A USING(SELECT ** AS **, ** AS **, ...) B
template <typename V>
std::wstring    make_using_part(wchar_t const*      schema_table_name,
                                std::wstring const& A_name  ,
                                std::wstring const& B_name  ,
                                safearrayRef&       attr    ,
                                V&&                 values  ,
                                bool                for_oracle)
{
    std::wstring prefix{ L"MERGE INTO " };
    (prefix += schema_table_name) += (L' ' + A_name + L" USING(SELECT ");
    valueof_variant vv;
    BSTR attributename;
    for (std::size_t j = 0; j < attr.getSize(1); ++j)
        if ( vv(std::forward<V>(values)(j), true) && (attributename = getBSTR(attr(j))) )
            prefix += vv.value() + L" AS " + attributename + L',';
    prefix.pop_back();
    if (for_oracle)     prefix += L" FROM DUAL) " + B_name;
    else                prefix += L") " + B_name;
    return prefix;
}

// ON (A.***=B.*** AND A.***=B.***)
std::wstring    make_ON_expr(std::wstring const&    A_name,
                             std::wstring const&    B_name,
                             safearrayRef&          attr,
                             safearrayRef&          key)
{ 
    std::wstring ret{ L" ON (" };
    std::size_t const attr_len = attr.getSize(1);
    std::size_t const key_len = key.getSize(1);
    std::size_t const bound = attr_len < key_len ? attr_len : key_len;
    auto varDest = iVariant();
    BSTR b;
    for (std::size_t i = 0; i < bound; ++i)
    {
        if (S_OK == ::VariantChangeType(&varDest, &key(i), 0, VT_I4)
            && 0 != varDest.lVal                    // キーであるカラム
            && (b = getBSTR(attr(i))))
        {
            ret += (A_name + L'.' + b + L'=' + B_name + L'.' + b + L" AND ");
        }
    }
    ret.erase(ret.end() - 5, ret.end());
    return ret + L')';
}

//  WHEN MATCHED THEN UPDATE SET *** = B.***, ...
std::wstring    make_set_part(std::wstring const& B_name, safearrayRef& attr, safearrayRef& key)
{
    std::wstring ret{ L" WHEN MATCHED THEN UPDATE SET " };
    std::size_t const key_len = key.getSize(1);
    std::size_t const bound = attr.getSize(1);
    auto varDest = iVariant();
    BSTR attributename;
    for (std::size_t i = 0; i < bound; ++i)
    {
        if ((key_len <= i ||
            (S_OK == ::VariantChangeType(&varDest, &key(i), 0, VT_I4) && 0 == varDest.lVal)        // キーでないカラム
             )
            && (attributename = getBSTR(attr(i))))
        {
            (ret += attributename) += (L'=' + B_name + L'.' + attributename + L',');
        }
    }
    if (!ret.empty())   ret.pop_back();
    return ret;
}

// WHEN NOT MATCHED THEN INSERT (***,***,...)  VALUES (B.***, B.***, ...);
std::wstring    make_values_part(std::wstring const& B_name, safearrayRef& attr)
{
    std::wstring ret1{ L" WHEN NOT MATCHED THEN INSERT (" };
    std::wstring ret2{ L" VALUES (" };
    std::size_t const bound = attr.getSize(1);
    BSTR attributename;
    for (std::size_t i = 0; i < bound; ++i)
        if (nullptr != (attributename = getBSTR(attr(i))))
        {
            (ret1 += attributename) += L',';
            ret2 += B_name + L'.' + attributename + L',';
        }
    ret1.pop_back();
    ret2.pop_back();
    return ret1 + L')' + ret2 + L");";
}

std::pair<std::wstring, std::wstring>
getTemporaryTableNames(std::wstring const& name)
{
    if (name.empty())                           return { L"A1", L"A2" };
    std::wstring::size_type p{ 0 };
    if (std::wstring::npos == (p = name.find_last_of(L'.')))    p = 0;
    else if (p + 1 < name.size())                               ++p;
    if (name[p] == L'A' || name[p] == L'a')     return { L"B1", L"B2" };
    else                                        return { L"A1", L"A2" };
}

}   //namespace mymd
