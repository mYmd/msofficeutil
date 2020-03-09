//TaskScheduler_vba.cpp
//Copyright (c) 2019 mmYYmmdd

#include "stdafx.h"
#include "msofficeutil.h"
#include "TaskScheduler.hpp"

// スケージュールされているタスクの名前と状態と引数を列挙
VARIANT __stdcall EnumTaskArguments(VARIANT const& Folder)
{
    auto fol = mymd::getBSTR(Folder);
    auto vec = mymd::enumTaskArguments(fol, true);
    if (vec.size())
    {
        std::vector<VARIANT> ret;
        std::wstring onoff[2] = {L"1", L"0"};
        for (auto& e: vec)
        {
            std::wstring* tmp[] = {&std::get<0>(e), std::get<1>(e)? onoff: onoff+1,  &std::get<2>(e)};
            auto elem = mymd::range2VArray(tmp, tmp+3, [](std::wstring* w) { return mymd::bstrVariant(*w); });
            ret.push_back(elem);
        }
        return mymd::range2VArray(ret.begin(), ret.end());
    }
    else
        return mymd::iVariant();
}

// タスクの引数を編集
__int32 __stdcall EditTaskArguments(VARIANT const& Folder, VARIANT const& TaskName, VARIANT const& Arguments)
{
    auto fol = mymd::getBSTR(Folder);
    auto taskName = mymd::getBSTR(TaskName);
    if(!fol || !taskName)                   return 0;
    mymd::safearrayRef ar(Arguments);
    if (ar.getDim() != 1 && 0 == ar.getSize(1))     return 0;
    std::vector<std::wstring> vec;
    for (std::size_t i = 0; i < ar.getSize(1); ++i)
        vec.emplace_back(mymd::getBSTR(ar(i)));
    return mymd::editTaskArguments(fol, taskName, &*vec.begin(), &*vec.end(), true)? 1: 0;
}

//  タスクの有効／無効状態を取得
__int32 __stdcall Task_ON_OFF_Status(VARIANT const& Folder, VARIANT const& TaskName)
{
    auto fol = mymd::getBSTR(Folder);
    auto taskName = mymd::getBSTR(TaskName);
    if(!fol || !taskName)                   return 0;
    return mymd::task_ON_OFF_status(fol, taskName, true)? 1: 0;
}

//  タスクの有効／無効を設定
__int32 __stdcall Task_ON_OFF_Change(VARIANT const& Folder, VARIANT const& TaskName, __int64 setOn)
{
    auto fol = mymd::getBSTR(Folder);
    auto taskName = mymd::getBSTR(TaskName);
    if(!fol || !taskName)                   return 0;
    return mymd::task_ON_OFF_Change(fol, taskName, 0!=setOn, true)? 1: 0;
}
