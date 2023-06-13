#pragma once

//#include "globals.hpp"
using Time = double;

#include <string>
#include <vector>
#include "matrix.h"

#ifdef __cplusplus
extern "C" {
#endif
    __declspec(dllexport)
        void __stdcall putBlackScholes(
            const double            spot,
            const double            vol,
            const bool              qSpot,
            const double            rate,
            const double            div,
            const std::string& store);

    __declspec(dllexport)
        void __stdcall putBarrier(
            const double            strike,
            const double            barrier,
            const Time              maturity,
            const double            monitorFreq,
            const double            smooth,
            const std::string& store);

    __declspec(dllexport)
        void __stdcall putDupire(
            const double            spot,
            const std::vector<double>& spots,
            const std::vector<Time>& times,
            //  spot major
            const matrix<double>& vols,
            const double            maxDt,
            const std::string& store);

#ifdef __cplusplus
}
#endif
