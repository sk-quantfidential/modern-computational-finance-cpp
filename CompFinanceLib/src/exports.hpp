#pragma once

#include "globals.hpp"
#include "matrix.h"
#include "main.h"
#include <string>
#include <vector>

template <class T>
class Model;

template <class T>
class Product;


//#ifdef __cplusplus
//extern "C" {
//#endif

    __declspec(dllexport)
        void /*__stdcall*/ restartThreadPool(
            double                  xNthread);

    __declspec(dllexport)
        void /*__stdcall*/ putBlackScholes(
            const double            spot,
            const double            vol,
            const bool              qSpot,
            const double            rate,
            const double            div,
            const std::string& store);

    __declspec(dllexport)
        void /*__stdcall*/ putDupire(
            const double            spot,
            const std::vector<double>& spots,
            const std::vector<Time>& times,
            //  spot major
            const matrix<double>&   vols,
            const double            maxDt,
            const std::string&      store);

    __declspec(dllexport)
        void /*__stdcall*/ putEuropean(
            const double            strike,
            const Time              exerciseDate,
            const Time              settlementDate,
            const std::string& store);

    __declspec(dllexport)
        void /*__stdcall*/ putBarrier(
            const double            strike,
            const double            barrier,
            const Time              maturity,
            const double            monitorFreq,
            const double            smooth,
            const std::string& store);

    __declspec(dllexport)
        void /*__stdcall*/ putContingent(
            const double            coupon,
            const Time              maturity,
            const double            payFreq,
            const double            smooth,
            const std::string& store);

    __declspec(dllexport)
        void /*__stdcall*/ putEuropeans(
            //  maturities must be given in increasing order
            const std::vector<double>& maturities,
            const std::vector<double>& strikes,
            const std::string&      store);

    __declspec(dllexport)
        Model<double>* /*__stdcall*/ getModel(
            const std::string&      modelId);

    __declspec(dllexport)
        Product<double>* /*__stdcall*/ getProduct(
            const std::string&      productId);

    __declspec(dllexport)
        ValueResults /*__stdcall*/ simValue(
            const std::string&      modelId,
            const std::string&      productId,
            const NumericalParam&   num);

    __declspec(dllexport)
        RiskPayoffResults /*__stdcall*/ AADriskOne(
            const std::string& modelId,
            const std::string& productId,
            const NumericalParam& num,
            const std::string& riskPayoff = "");

    __declspec(dllexport)
        RiskPayoffResults /*__stdcall*/ AADriskAggregate(
            const std::string& modelId,
            const std::string& productId,
            const std::map<std::string, double>& notionals,
            const NumericalParam& num);

    __declspec(dllexport)
        RiskReports /*__stdcall*/ bumpRisk(
            const std::string& modelId,
            const std::string& productId,
            const NumericalParam& num);

    __declspec(dllexport)
        RiskReports /*__stdcall*/ AADriskMulti(
            const std::string& modelId,
            const std::string& productId,
            const NumericalParam& num);

    __declspec(dllexport)
        RiskReports* /*__stdcall*/ findRiskReports(
            const std::string& riskId);

    __declspec(dllexport)
        DupireCalibResults /*__stdcall*/ dupireCalib(
            const std::vector<double>& inclSpots,
            const double maxDs,
            const std::vector<Time>& inclTimes,
            const double maxDt,
            const double spot,
            const double vol,
            const double jmpIntens = 0.0,
            const double jmpAverage = 0.0,
            const double jmpStd = 0.0);

    __declspec(dllexport)
        SuperbucketResults /*__stdcall*/ dupireSuperbucket(
            const double            spot,
            const double            maxDt,
            const std::string& productId,
            const map<string, double>& notionals,
            const std::vector<double>& inclSpots,
            const double            maxDs,
            const std::vector<Time>& inclTimes,
            const double            maxDtVol,
            const std::vector<double>& strikes,
            const std::vector<Time>& mats,
            const double            vol,
            const double            jmpIntens,
            const double            jmpAverage,
            const double            jmpStd,
            const NumericalParam& num);
 
     __declspec(dllexport)
         SuperbucketResults /*__stdcall*/ dupireSuperbucketBump(
            const double            spot,
            const double            maxDt,
            const std::string& productId,
            const map<string, double>& notionals,
            const std::vector<double>& inclSpots,
            const double            maxDs,
            const std::vector<Time>& inclTimes,
            const double            maxDtVol,
            const std::vector<double>& strikes,
            const std::vector<Time>& mats,
            const double            vol,
            const double            jmpIntens,
            const double            jmpAverage,
            const double            jmpStd,
            const NumericalParam& num);

    __declspec(dllexport)
        double /*__stdcall*/  merton(
            const double            spot,
            const double            strike,
            const double            vol,
            const Time              mat,
            const double            intens,
            const double            meanJmp,
            const double            stdJmp);

    __declspec(dllexport)
        double /*__stdcall*/ blackScholesKO(
            const double spot,
            const double rate,
            const double div,
            const double strike,
            const double barrier,
            const double mat,
            const double vol);

//#ifdef __cplusplus
//}
//#endif
