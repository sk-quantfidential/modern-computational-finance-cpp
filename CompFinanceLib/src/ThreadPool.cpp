/*
Written by Antoine Savine in 2018

This code is the strict IP of Antoine Savine

License to use and alter this code for personal and commercial applications
is freely granted to any person or company who purchased a copy of the book

Modern Computational Finance: AAD and Parallel Simulations
Antoine Savine
Wiley, 2018

As long as this comment is preserved at the top of the file
*/

#include "ThreadPool.h"
#include "exports.hpp"

__declspec(dllexport)
void __stdcall restartThreadPool(
    const int   numThread)
{
    // Todo: restartThreadPool(numThread);
    ThreadPool::getInstance()->stop();
    ThreadPool::getInstance()->start(numThread);
}

//  Statics
ThreadPool ThreadPool::myInstance;
thread_local size_t ThreadPool::myTLSNum = 0;