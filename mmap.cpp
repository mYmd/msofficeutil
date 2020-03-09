    //Copyright (c) 2020 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"

#define PICOJSON_USE_INT64 1
#include "picojson.h"

//#undef GetFreeSpace
#import "scrrun.dll"  raw_interfaces_only, raw_native_types, named_guids, \
        rename("GetFreeSpace", "_GetFreeSpace"),rename("DeleteFile", "_DeleteFile"), \
        rename("DeleteFile", "_DeleteFile"), rename("MoveFile", "_MoveFile"), rename("CopyFile", "_CopyFile")
//scrrun.tlh

#if  defined _M_X64
using vbLongPtr_t = __int64;
#else
using vbLongPtr_t = __int32;
#endif

namespace   {

    std::string WideCharToMultiByte_(VARIANT const& v)
    {
        auto bstr = mymd::getBSTR(v);
        if (!bstr)  return std::string{};
        auto n = ::WideCharToMultiByte(CP_UTF8, 0, bstr, -1, nullptr, 0, nullptr, nullptr);
        std::string buf(n-1, '\0');
        ::WideCharToMultiByte(CP_UTF8, 0, bstr, -1, &buf[0], n, nullptr, nullptr);
        return buf;
    }


    BSTR MultiByteToWideChar_(std::string const& s)
    {
        auto n = ::MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), nullptr, 0);
        auto bstr = ::SysAllocStringByteLen(NULL, n * sizeof(wchar_t));
        ::MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), bstr, n);
        return bstr;
    }

            
    picojson::value elem2pico(VARIANT const& v)
    {
        switch (v.vt)
        {
        case VT_EMPTY:  case VT_NULL:
            return picojson::value{};
        case VT_ERROR:
            return picojson::value{static_cast<int64_t>(v.scode)};
        case VT_BOOL:
            return picojson::value{v.boolVal? true: false};
        case VT_I2:
            return picojson::value{static_cast<int64_t>(v.iVal)};
        case VT_I4:
            return picojson::value{static_cast<int64_t>(v.lVal)};
        case VT_I8:
            return picojson::value{static_cast<int64_t>(v.llVal)};
        case VT_DATE:
            return picojson::value{v.date};
        case VT_R4:
            return picojson::value{v.fltVal};
        case VT_R8:
            return picojson::value{v.dblVal};
        case VT_BSTR:
            return picojson::value{WideCharToMultiByte_(v)};
        default:
            auto tmp = mymd::iVariant();
            ::VariantChangeType(&tmp, &v, 0, VT_BSTR);
            auto ret = elem2pico(tmp);
            ::VariantClear(&tmp);
            return ret;
        }
    }

    picojson::value variant2pico(VARIANT const&);

    picojson::value dictionary2pico(Scripting::IDictionaryPtr pDic)
    {
        picojson::object map;
        VARIANT Key = mymd::iVariant();
        VARIANT Item = mymd::iVariant();
        pDic->Keys(&Key);
        pDic->Items(&Item);
        mymd::safearrayRef refK(Key);
        mymd::safearrayRef refI(Item);
        for ( std::size_t i = 0; i < refK.getSize(1); ++i )
        {
            ::VariantChangeType(&refK(i), &refK(i), 0, VT_BSTR);
            map.emplace(WideCharToMultiByte_(refK(i)), variant2pico(refI(i)));
        }
        return picojson::value(std::move(map));
    }


    picojson::value variant2pico(VARIANT const& data)
    {
        mymd::safearrayRef ref{data};
        switch (ref.getDim())
        {
        case 0:
        {
            Scripting::IDictionaryPtr pDic{nullptr};
            long count{0};
            if ( data.vt == VT_DISPATCH && S_OK == data.pdispVal->QueryInterface(Scripting::IID_IDictionary, reinterpret_cast<void**>(&pDic)) )
                      // ### Scripting.Dictionaryの場合
                return dictionary2pico(pDic);
            else
                return elem2pico(data);
        }
        case 1:
        {
            picojson::array arr;
            arr.reserve(ref.getSize(1));
            for ( std::size_t i = 0; i < ref.getSize(1); ++i )
                arr.push_back(variant2pico(ref(i)));
            return picojson::value(std::move(arr));
        }
        case 2:
        {
            picojson::array arr;
            arr.reserve(ref.getSize(1));
            for ( std::size_t i = 0; i < ref.getSize(1); ++i )
            {
                picojson::array arr2;
                arr2.reserve(ref.getSize(2));
                for ( std::size_t j = 0; j < ref.getSize(2); ++j )
                    arr2.push_back(variant2pico(ref(i, j)));
                arr.push_back(picojson::value(std::move(arr2)));
            }
            return picojson::value(std::move(arr));
        }
        default:
            return picojson::value{};
        }
    }


    VARIANT pico2variant(picojson::value const&);

    VARIANT dictionary2variant(picojson::object const& dict)
    {
        Scripting::IDictionary* pDictionary;
        if (S_OK != ::CoCreateInstance(Scripting::CLSID_Dictionary, NULL, CLSCTX_INPROC_SERVER, Scripting::IID_IDictionary, reinterpret_cast<LPVOID*>(&pDictionary)))
            return mymd::iVariant();
        for ( auto const& e : dict )
        {
            auto key = mymd::bstrVariant(MultiByteToWideChar_(e.first));
            auto elem = pico2variant(e.second);
            pDictionary->Add(&key, &elem);
            ::VariantClear(&key);
            ::VariantClear(&elem);
        }
        auto v = mymd::iVariant(VT_DISPATCH);
        v.pdispVal = pDictionary;
        return v;
    }


    VARIANT pico2variant(picojson::value const& data)
    {
        if(data.is<bool>())
        {
            auto v = mymd::iVariant(VT_BOOL);
            v.boolVal = data.get<bool>()? VARIANT_TRUE: VARIANT_FALSE;
            return v;
        }
        else if (data.is<int64_t>())
        {
            auto val = data.get<int64_t>();
            if ( LONG_MIN <= val && val <= LONG_MAX )
            {
                auto v = mymd::iVariant(VT_I4);
                v.lVal = static_cast<__int32>(val);
                return v;
            }
            else
            {
                auto v = mymd::iVariant(VT_I8);
                v.llVal = val;
                return v;
            }
        }
        else if (data.is<std::string>())
        {
            std::string s{data.get<std::string>()};
            auto bstr = MultiByteToWideChar_(s);
            auto v = mymd::iVariant(VT_BSTR);
            v.bstrVal = bstr;
            return v;
        }
        else if (data.is<double>())
        {
            auto v = mymd::iVariant(VT_R8);
            v.dblVal = data.get<double>();
            return v;
        }
        else if (data.is<picojson::array>())
        {
            auto const& refv = data.get<picojson::array>();
            std::vector<VARIANT> vec;
            vec.reserve(refv.size());
            for ( std::size_t i = 0; i < refv.size(); ++i )
            {
                vec.push_back(pico2variant(refv[i]));
            }
            return mymd::range2VArray(vec.begin(), vec.end());
        }
        else if (data.is<picojson::object>())   //### 
            return dictionary2variant(data.get<picojson::object>());
        else
            return mymd::iVariant();
    }
}


vbLongPtr_t __stdcall
write_mmap(VARIANT const& data_, VARIANT const& name_, __int32& bytes_)
{
    auto name = mymd::getBSTR(name_);
    if (!name)           return 0;
    auto pico = variant2pico(data_);
    auto str = pico.serialize(false);
    bytes_ = static_cast<__int32>(str.length());
    HANDLE ret = ::CreateFileMappingW(INVALID_HANDLE_VALUE,
                                      NULL,
                                      PAGE_READWRITE,
                                      0,
                                      static_cast<DWORD>(str.length()+1),
                                      name);
    if (!ret)                   return 0;
    auto p = static_cast<char*>(::MapViewOfFile(ret, FILE_MAP_WRITE, 0, 0, str.length()+1));
    if (!p)                     return 0;
    ::strcpy_s(p, str.length()+1, str.data());
    ::UnmapViewOfFile(p);
    //::CloseHandle(ret);
    return reinterpret_cast<vbLongPtr_t>(ret);
}


vbLongPtr_t __stdcall
reserve_mmap(__int32 bytes_, VARIANT const& name_)
{
    if (bytes_ <= 0)        return 0;
    auto name = mymd::getBSTR(name_);
    if (!name )             return 0;
    HANDLE ret = ::CreateFileMappingW(INVALID_HANDLE_VALUE,
                                      NULL,
                                      PAGE_READWRITE,
                                      0,
                                      bytes_,
                                      name);
    return reinterpret_cast<vbLongPtr_t>(ret);
}


void __stdcall
closeHandle(vbLongPtr_t h)
{
    ::CloseHandle(reinterpret_cast<HANDLE>(h));
}


void __stdcall
read_mmap(VARIANT const& name_, VARIANT& out, __int32 bytes)
{
    ::VariantClear(&out);
    auto name = mymd::getBSTR(name_);
    if ( !name )                return;
    auto hSharedMemory = ::OpenFileMapping(FILE_MAP_READ, TRUE, name);
    if ( !hSharedMemory )       return;
    // 共有メモリのマッピング
    auto cpSheredMemory = static_cast<char*>(::MapViewOfFile(hSharedMemory, FILE_MAP_READ, 0, 0, bytes));
    if ( !cpSheredMemory )      return;
    picojson::value pico;
    picojson::parse(pico, std::string{cpSheredMemory});
    out = pico2variant(pico);
    ::UnmapViewOfFile(cpSheredMemory);
    ::CloseHandle(hSharedMemory);
}


VARIANT __stdcall
VariantToJson(VARIANT const& data_)
{
    auto pico = variant2pico(data_);
    auto str = pico.serialize(false);
    auto bstr = MultiByteToWideChar_(str);
    auto v = mymd::iVariant(VT_BSTR);
    v.bstrVal = bstr;
    return v;
}


void __stdcall
JsonToVariant(VARIANT const& data_, VARIANT& out)
{
    ::VariantClear(&out);
    std::string buf = WideCharToMultiByte_(data_);
    picojson::value pico;
    picojson::parse(pico, buf);
    out = pico2variant(pico);
}
