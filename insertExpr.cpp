//insertExpr.cpp
//Copyright (c) 2018 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include <vector>
#include <limits>

namespace   {

    // VARIANT の各種値を
    class   valueof_variant {
        std::wstring value_;
    public:
        std::wstring const& value() const { return value_; }
        bool operator ()(VARIANT const& v, bool evaluate_null = false);
    };

    template <typename V>
    std::wstring  make_insert_expr_imple(valueof_variant&, std::wstring, mymd::safearrayRef&, V&&, bool);

    std::wstring  make_bulk_insert_attr_part(mymd::safearrayRef&);

    template <typename V>
    std::wstring  make_bulk_insert_value_part(valueof_variant&, V&&, std::size_t);

    // MERGE INTO **** A USING(SELECT ** AS **, ** AS **, ...) B
    template <typename V>
    std::wstring  make_using_part(valueof_variant&      vv,
                                  wchar_t const*        schema_table_name,
                                  std::wstring const&   alias1  ,
                                  std::wstring const&   alias2  ,
                                  mymd::safearrayRef&   attr    ,
                                  V&&                   values  ,
                                  bool                  for_oracle);

    // ON (A.***=B.*** AND A.***=B.***)
    std::wstring  make_ON_part(std::wstring const& alias1, std::wstring const& alias2, mymd::safearrayRef& attr, mymd::safearrayRef& key);

    //  WHEN MATCHED THEN UPDATE SET *** = B.***, ...
    std::wstring  make_set_part(std::wstring const& alias2, mymd::safearrayRef& attr, mymd::safearrayRef& key);

    // WHEN NOT MATCHED THEN INSERT (***,***,...)  VALUES (B.***, B.***, ...);
    std::wstring  make_values_part(std::wstring const& alias2, mymd::safearrayRef& attr);

    std::pair<std::wstring, std::wstring> const&  getAliasTableNames(std::wstring const& name);

}

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
    valueof_variant  vv;
    switch (values.getDim())
    {
    case 1:
    {
        if (0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt))    //単一レコード
        {
            auto const bound = values.getSize(1);
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(j) : dummy;
            };
            return mymd::bstrVariant(make_insert_expr_imple(vv, prefix, attr, value_func, true));
        }
        else            //複数レコードのジャグ配列
        {
            for (std::size_t i = 0; i < values.getSize(1); ++i)
            {
                mymd::safearrayRef value2{ values(i) };
                auto const bound = value2.getSize(1);
                auto value_func = [&](std::size_t j) ->VARIANT const& {
                    return (j < bound) ? value2(j) : dummy;
                };
                buf.push_back(make_insert_expr_imple(vv, prefix, attr, value_func, true));
            }
        }
        break;
    }
    case 2:            //複数レコードの2次元配列
    {
        auto const bound = values.getSize(2);
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(make_insert_expr_imple(vv, prefix, attr, value_func, true));
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

// MERGE文の生成
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
    auto const altname = getAliasTableNames(schema_table_name);
    std::wstring const& alias1 = altname.first;
    std::wstring const& alias2 = altname.second;
    auto dummy = mymd::iVariant();
    valueof_variant vv;
    auto for_oracle = (0 != oracle);
    switch (values.getDim())
    {
    case 1:
    {
        if (0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt))    //単一レコード
        {
            auto const bound = values.getSize(1);
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(j) : dummy;
            };
            return mymd::bstrVariant( make_using_part(vv, schema_table_name, alias1, alias2, attr, value_func, for_oracle)
                                    + make_ON_part(alias1, alias2, attr, key)
                                    + make_set_part(alias2, attr, key)
                                    + make_values_part(alias2, attr)
                                    );
        }
        else            //複数レコードのジャグ配列
        {
            for (std::size_t i = 0; i < values.getSize(1); ++i)
            {
                mymd::safearrayRef value2{ values(i) };
                auto const bound = value2.getSize(1);
                auto value_func = [&](std::size_t j) ->VARIANT const& {
                    return (j < bound) ? value2(j) : dummy;
                };
                buf.push_back(  make_using_part(vv, schema_table_name, alias1, alias2, attr, value_func, for_oracle)
                              + make_ON_part(alias1, alias2, attr, key)
                              + make_set_part(alias2, attr, key)
                              + make_values_part(alias2, attr)
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
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(  make_using_part(vv, schema_table_name, alias1, alias2, attr, value_func, for_oracle)
                          + make_ON_part(alias1, alias2, attr, key)
                          + make_set_part(alias2, attr, key)
                          + make_values_part(alias2, attr)
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

// BULK_INSERT文の生成
VARIANT __stdcall
make_bulk_insert_expr(  VARIANT const&  schema_table_name_,
                        VARIANT const&  attr_   ,
                        VARIANT const&  values_ ,
                        __int32         unit_size_)
{
    auto schema_table_name = mymd::getBSTR(schema_table_name_);
    if (!schema_table_name)         return mymd::iVariant();
    mymd::safearrayRef attr{ attr_ };
    mymd::safearrayRef values{ values_ };
    if (attr.getDim() == 0 || values.getDim() == 0)     return mymd::iVariant();
    if ( values.getDim() == 1 && 0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt) )    //単一レコード
        return make_insert_expr(schema_table_name_, attr_, values_);
    //
    std::wstring  expr{ L"INSERT INTO " };
    expr += schema_table_name;
    auto attr_part = make_bulk_insert_attr_part(attr);
    if (attr_part.empty())        return mymd::iVariant();
    expr += attr_part + L" VALUES ";
    std::vector<std::wstring>   buf;
    auto dummy = mymd::iVariant();
    valueof_variant  vv;
    std::size_t const attrSize = attr.getSize(1);
    switch (values.getDim())
    {
    case 1:
    {
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            mymd::safearrayRef value2{ values(i) };
            auto const bound = value2.getSize(1);
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? value2(j) : dummy;
            };
            buf.push_back(make_bulk_insert_value_part(vv, value_func, attrSize));
        }
        break;
    }
    case 2:            //複数レコードの2次元配列
    {
        auto const bound = values.getSize(2);
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(make_bulk_insert_value_part(vv, value_func, attrSize));
        }
        break;
    }
    default:
        return mymd::iVariant();
    }
    auto std_min = [](std::size_t a, std::size_t b) { return a <= b ? a : b; };
    auto const unit_size = std_min(values.getSize(1), static_cast<std::size_t>(0 <= unit_size_ ? unit_size_ : 256));
    auto const ret_size = (values.getSize(1) + unit_size - 1) / unit_size;
    std::vector<VARIANT> ret;
    for (std::size_t i = 0; i < ret_size; ++i)
    {
        auto tmp = expr;
        std::for_each(buf.begin() + (i * unit_size),
                      buf.begin() + std_min(buf.size(), (i + 1) * unit_size),
                      [&](std::wstring const& w) { tmp += w; });
        tmp.pop_back();
        tmp += L';';
        ret.push_back(mymd::bstrVariant(tmp));
    }
    return mymd::range2VArray(ret.begin(), ret.end());
}

// ORACLE の INSERT_ALL 文の生成
VARIANT __stdcall
make_insert_all_expr(VARIANT const&  schema_table_name_,
                     VARIANT const&  attr_,
                     VARIANT const&  values_,
                     __int32         unit_size_)
{
    auto schema_table_name = mymd::getBSTR(schema_table_name_);
    if (!schema_table_name)         return mymd::iVariant();
    mymd::safearrayRef attr{ attr_ };
    mymd::safearrayRef values{ values_ };
    if (attr.getDim() == 0 || values.getDim() == 0)       return mymd::iVariant();
    std::wstring prefix{ L"INTO " };
    (prefix += schema_table_name) += L" (";
    std::vector<std::wstring>   buf;
    auto dummy = mymd::iVariant();
    valueof_variant vv;
    switch (values.getDim())
    {
    case 1:
    {
        if (0 < values.getSize(1) && 0 == (VT_ARRAY & values(0).vt))    //単一レコード
        {
            auto const bound = values.getSize(1);
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(j) : dummy;
            };
            return mymd::bstrVariant(make_insert_expr_imple(vv, L"INSERT " + prefix, attr, value_func, false));
        }
        else            //複数レコードのジャグ配列
        {
            for (std::size_t i = 0; i < values.getSize(1); ++i)
            {
                mymd::safearrayRef value2{ values(i) };
                auto const bound = value2.getSize(1);
                auto value_func = [&](std::size_t j) ->VARIANT const& {
                    return (j < bound) ? value2(j) : dummy;
                };
                buf.push_back(make_insert_expr_imple(vv, prefix, attr, value_func, false));
            }
        }
        break;
    }
    case 2:            //複数レコードの2次元配列
    {
        auto const bound = values.getSize(2);
        for (std::size_t i = 0; i < values.getSize(1); ++i)
        {
            auto value_func = [&](std::size_t j) ->VARIANT const& {
                return (j < bound) ? values(i, j) : dummy;
            };
            buf.push_back(make_insert_expr_imple(vv, prefix, attr, value_func, false));
        }
        break;
    }
    default:
        return mymd::iVariant();
    }
    auto std_min = [](std::size_t a, std::size_t b) { return a <= b ? a : b; };
    auto const unit_size = std_min(values.getSize(1), static_cast<std::size_t>(0 <= unit_size_ ? unit_size_ : 256));
    auto const ret_size = (values.getSize(1) + unit_size - 1) / unit_size;
    std::vector<VARIANT> ret;
    for (std::size_t i = 0; i < ret_size; ++i)
    {
        std::wstring tmp{ L"INSERT ALL " };
        std::for_each(buf.begin() + (i * unit_size),
                        buf.begin() + std_min(buf.size(), (i + 1) * unit_size),
                        [&](std::wstring const& w) { tmp += w; });
        tmp += L"SELECT * FROM DUAL;";
        ret.push_back(mymd::bstrVariant(tmp));
    }
    return mymd::range2VArray(ret.begin(), ret.end());
}

namespace   {

template <typename V>
std::wstring    make_insert_expr_imple(valueof_variant&     vv      ,
                                       std::wstring         prefix  ,
                                       mymd::safearrayRef&  attr    ,
                                       V&&                  values  ,
                                       bool                 semicollon)
{
    std::wstring value_part{ L") VALUES (" };
    bool valid{ false };
    for (std::size_t j = 0; j < attr.getSize(1); ++j)
    {
        BSTR attributename;
        if ( vv(std::forward<V>(values)(j)) && (attributename = mymd::getBSTR(attr(j))) )
        {
            prefix += attributename;
            prefix += L',';
            value_part += vv.value() + L',';
            valid = true;
        }
    }
    if (valid)
    {
        prefix.pop_back();
        value_part.pop_back();
        return prefix + value_part + (semicollon ? L");" : L") ");
    }
    else
    {
        return std::wstring{};
    }
}

//  (F1,F2,...FN)
std::wstring    make_bulk_insert_attr_part(mymd::safearrayRef&  attr)
{
    std::wstring attr_part{ L" (" };
    for (std::size_t j = 0; j < attr.getSize(1); ++j)
    {
        auto b = mymd::getBSTR(attr(j));
        if (!b)   return std::wstring{};
        (attr_part += b) += L',';
    }
    attr_part.pop_back();
    return attr_part + L')';
}

// (value1,value2,..,valueN)   + (,| )
template <typename V>
std::wstring    make_bulk_insert_value_part(valueof_variant& vv, V&& values, std::size_t attrSize)
{
    std::wstring value_part{ L'(' };
    for ( std::size_t j = 0; j < attrSize; ++j )
    {
        if ( vv(std::forward<V>(values)(j)) )   value_part += vv.value() + L',';
        else                                    value_part += L"NULL,";
    }
    value_part.pop_back();
    return value_part + L"),";
}

//***********************************************************************
// MERGE INTO **** A USING(SELECT ** AS **, ** AS **, ...) B
template <typename V>
std::wstring    make_using_part(valueof_variant&    vv      ,
                                wchar_t const*      schema_table_name,
                                std::wstring const& alias1  ,
                                std::wstring const& alias2  ,
                                mymd::safearrayRef& attr    ,
                                V&&                 values  ,
                                bool                for_oracle)
{
    std::wstring prefix{ L"MERGE INTO " };
    (prefix += schema_table_name) += (L' ' + alias1 + L" USING(SELECT ");
    for (std::size_t j = 0; j < attr.getSize(1); ++j)
    {
        BSTR attributename;
        if (vv(std::forward<V>(values)(j), true) && (attributename = mymd::getBSTR(attr(j))))
            prefix += vv.value() + L" AS " + attributename + L',';
    }
    prefix.pop_back();
    if (for_oracle)     prefix += L" FROM DUAL) " + alias2;
    else                prefix += L") " + alias2;
    return prefix;
}

// ON (A.***=B.*** AND A.***=B.***)
std::wstring    make_ON_part(std::wstring const&    alias1  ,
                             std::wstring const&    alias2  ,
                             mymd::safearrayRef&    attr,
                             mymd::safearrayRef&    key)
{ 
    std::wstring ret{ L" ON (" };
    std::size_t const attr_len = attr.getSize(1);
    std::size_t const key_len = key.getSize(1);
    std::size_t const bound = attr_len < key_len ? attr_len : key_len;
    auto varDest = mymd::iVariant();
    for (std::size_t i = 0; i < bound; ++i)
    {
        BSTR attributename;
        if ( S_OK == ::VariantChangeType(&varDest, &key(i), 0, VT_I4)
             && 0 != varDest.lVal                    // キーであるカラム
             && (attributename = mymd::getBSTR(attr(i))) )
        {
            ret += (alias1 + L'.' + attributename + L'=' + alias2 + L'.' + attributename + L" AND ");
        }
    }
    ret.erase(ret.end() - 5, ret.end());
    return ret + L')';
}

//  WHEN MATCHED THEN UPDATE SET *** = B.***, ...
std::wstring    make_set_part(std::wstring const& alias2, mymd::safearrayRef& attr, mymd::safearrayRef& key)
{
    std::wstring ret{ L" WHEN MATCHED THEN UPDATE SET " };
    std::size_t const key_len = key.getSize(1);
    std::size_t const bound = attr.getSize(1);
    auto varDest = mymd::iVariant();
    for (std::size_t i = 0; i < bound; ++i)
    {
        BSTR attributename;
        if ((key_len <= i ||
            (S_OK == ::VariantChangeType(&varDest, &key(i), 0, VT_I4) && 0 == varDest.lVal)        // キーでないカラム
             )
            && (attributename = mymd::getBSTR(attr(i))))
        {
            (ret += attributename) += (L'=' + alias2 + L'.' + attributename + L',');
        }
    }
    if (!ret.empty())   ret.pop_back();
    return ret;
}

// WHEN NOT MATCHED THEN INSERT (***,***,...)  VALUES (B.***, B.***, ...);
std::wstring    make_values_part(std::wstring const& alias2, mymd::safearrayRef& attr)
{
    std::wstring ret1{ L" WHEN NOT MATCHED THEN INSERT (" };
    std::wstring ret2{ L" VALUES (" };
    std::size_t const bound = attr.getSize(1);
    for (std::size_t i = 0; i < bound; ++i)
    {
        BSTR attributename;
        if (nullptr != (attributename = mymd::getBSTR(attr(i))))
        {
            (ret1 += attributename) += L',';
            ret2 += alias2 + L'.' + attributename + L',';
        }
    }
    ret1.pop_back();
    ret2.pop_back();
    return ret1 + L')' + ret2 + L");";
}

std::pair<std::wstring, std::wstring> const&
getAliasTableNames(std::wstring const& name)
{
    static std::pair<std::wstring, std::wstring> const A_Name{ L"A1", L"A2" };
    static std::pair<std::wstring, std::wstring> const B_Name{ L"B1", L"B2" };
    if (name.empty())                           return A_Name;
    std::wstring::size_type p{ 0 };
    if (std::wstring::npos == (p = name.find_last_of(L'.')))    p = 0;
    else if (p + 1 < name.size())                               ++p;
    if (name[p] == L'A' || name[p] == L'a')     return B_Name;
    else                                        return A_Name;
}

// *****************************************

// ' があれば重ねる
void  quote_double(std::wstring& w, BSTR s)
{
    w = L'x';
    w += s;
    std::wstring::size_type pos{ std::wstring::npos };
    while (std::wstring::npos != (pos = w.find_last_of(L'\'', pos)))    // 'があれば
    {
        w.insert(pos, 1, L'\'');
        if (0 == pos--) break;
    }
    w[0] = L'\'';
    w.push_back(L'\'');
}

// 浮動小数を文字列化
void to_wstring2(std::wstring& w, double val)
{
    //                       -   1.                 mantissa                 e   +  nnn \0
    const std::size_t size = 1 + 2 + std::numeric_limits<double>::digits10 + 1 + 1 + 3 + 1;
    w.resize(size);
    auto n = std::swprintf(&w[0], size, L"%.15e", val);
    w.resize(n);
}

bool valueof_variant::operator ()(VARIANT const& v, bool evaluate_null)
{
    value_.clear();
    auto b = mymd::getBSTR(v);
    if (b)
    {
        quote_double(value_, b);
        return true;
    }
    else
    {
        auto varDest = mymd::iVariant();
        switch (v.vt)
        {
        case VT_EMPTY: case VT_NULL:
            if (evaluate_null)      value_ = L"NULL";
            return  evaluate_null;
        case VT_R4:
            to_wstring2(value_, v.fltVal);
            return true;
        case VT_R8:
            to_wstring2(value_, v.dblVal);
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

}
