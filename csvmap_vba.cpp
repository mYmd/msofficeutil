//csvmap_vba.cpp
//Copyright (c) 2019 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include <vector>
#include "csvmap.hpp"

//CSV�t�@�C���̓ǂݍ���
// stream : ���̓X�g���[���ifstream �܂��� stringstream ��z�� �j
// R : �o�͂��镶���^�iwchar_t��z��j
// S : �Ώ�stream�̕����^�i�t�@�C����codeconvert����O��Ȃ�string��z��j
// delimiter : ��؂蕶���i����ɂ���Č^R �����肳���j
// elem_func   : �v�f���̃R�[���o�b�N�@�@[&](std::size_t count, std::wstring&& expr) ���A (���ڔԍ�, �v�f������)
// record_func : 1���R�[�h�ǏI���̃R�[���o�b�N�@[&](std::size_t count, std::size_t size)�@���A(�s�ԍ�, ���Y�s�̍��ڐ�)
// codepage : CP_UTF8, CP_ACP �Ȃ�
// �Ԃ�l : �ǂ񂾍s���i�e�L�X�g�t�@�C���̏ꍇ�͍Ō�̋�s���J�E���g�j

VARIANT __stdcall Map_CSV(VARIANT const&    Path        ,
                          VARIANT const&    Delimiter   ,
                          __int32 const     Code_Page   ,
                          __int32 const     Head_N      ,
                          __int32 const     Head_Cut    )
{
    auto path = mymd::getBSTR(Path);
    auto delimiter= mymd::getBSTR(Delimiter);
    if (!path || !delimiter)    return mymd::iVariant();
    std::ifstream ifs(path);
    if (ifs.bad())              return mymd::iVariant();
    std::vector<std::wstring> record;
    std::vector<VARIANT> vec;
    auto elem_func = [&](std::size_t count, std::wstring&& expr)    {
        record.push_back(std::move(expr));
    };
    auto head_cut = static_cast<std::size_t>(Head_Cut);
    auto head_n = static_cast<std::size_t>(Head_N < 0? INT_MAX: Head_N);
    auto record_func = [&](std::size_t count, std::size_t size) {
        if (head_cut <= count && count < head_n)
        {
            auto elem = mymd::range2VArray(record.begin(), record.end(), [](std::wstring& w) { return mymd::bstrVariant(w); });
            vec.push_back(elem);
            record.clear();
            return true;
        }
        return false;
    };
    auto s = mymd::map_csv(ifs, delimiter[0], elem_func, record_func, Code_Page);
    return mymd::range2VArray(vec.begin(), vec.end());
}
