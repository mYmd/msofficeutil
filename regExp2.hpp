//Copyright (c) 2018 mmYYmmdd
#pragma once
#include "msofficeutil.h"
#include <regex>
#include <vector>

namespace mymd  {

    //�m���ǂ݁A�ے��ǂ݂������������K�\��
    //������ (?<= �� (?<! �̓p�^�[���̐擪�Ɍ���
    class regexp2   {
        int positive_negative;
        std::wstring p_ahead, p_behind;
        std::wstring match_, prefix_, suffix_;
        std::vector<std::wstring> submatches_;
    public:
        regexp2(std::wstring const& pattern);
        void execute(std::wstring const& target, bool icase);
        std::wstring const& match() const;
        std::wstring const& prefix() const;
        std::wstring const& suffix() const;
        std::vector<std::wstring> const& submatches() const;
    };

}