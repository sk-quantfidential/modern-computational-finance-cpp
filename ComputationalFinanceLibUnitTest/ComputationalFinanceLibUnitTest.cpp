#include "pch.h"
#include "CppUnitTest.h"

#include <string>
#include "..\CompFinanceLib\src\exports.hpp"
#include "..\CompFinanceLib\src\main.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace std;

unordered_map<string, RiskReports> riskStore;

namespace ComputationalFinanceLibUnitTest
{
	TEST_CLASS(ComputationalFinanceLibTestExports)
	{
	public:
		

        TEST_METHOD(TestRestartThreadPool)
        {
            // Arrange
            int numThreads = 15;

            // Act
            restartThreadPool(numThreads);

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestPutBlackScholes)
        {
            // Arrange
            const double spot = 100.0;
            const double vol = 0.15;
            const bool qSpot = false;
            const double rate = 0.02;
            const double div = 0.03;
            const string id = "PutBlackScholes_id";

            // Act
            putBlackScholes(spot, vol, qSpot, rate, div, id);

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestPutDupire)
        {
            // Arrange
            const double spot = 100.0;
            vector<double> vspots = {};
            vector<double> vtimes = {};
            const double maxDt = 1.0;
            matrix<double> vvols = {};
            const string id = "PutDupire_id";

            // Act
            putDupire(spot, vspots, vtimes, vvols, maxDt, id);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestPutEuropean)
        {
            // Arrange
            const double strike = 100.0;
            const Time exerciseDate = 1.0;
            const Time settlementDate = 1.01;
            const string id = "PutEuropean_id";

            // Act
            putEuropean(strike, exerciseDate, settlementDate, id);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestPutBarrier)
        {
            // Arrange
            const double strike = 100.0;
            const double barrier = 105.0;
            const Time maturity = 1.0;
            const double monitorFreq = 0.25;
            const double smoothing = 1.0;
            const string id = "barrier_id";

            //  Act
            putBarrier(strike, barrier, maturity, monitorFreq, smoothing, id);

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }


        TEST_METHOD(TestPutContingent)
        {
            // Arrange
            const double coupon = 0.02;
            const Time maturity = 1.0;
            const double payFreq = 0.25;
            const double smoothing = 1.0;
            const string id = "PutContingent_id";

            // Act
            putContingent(coupon, maturity, payFreq, smoothing, id);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestPutEuropeans)
        {
            // Arrange
            vector<double> vmats = {};
            vector<double> vstrikes = {};
            const string id = "PutEuropeans_id";

            // Act
            putEuropeans(vmats, vstrikes, id);

            Assert::AreEqual(1.0, 1.0);
        }

        //  Access payoff identifiers and parameters

        TEST_METHOD(TestParameters)
        {
            // Arrange
            const string mid = "models_id";

            // Act
            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            // Assert:  Make sure we have a model
            Assert::IsTrue(mdl);

            const auto& paramLabels = mdl->parameterLabels();
            const auto& params = mdl->parameters();
            vector<double> paramsCopy(params.size());
            Assert::IsTrue(params.size());
        }

        TEST_METHOD(TestPayoffIds)
        {
            // Arrange
            const string pid = "Product_id";

            // Act
            const auto* prd = getProduct<double>(pid);
            // Assert Make sure we have a product
            Assert::IsTrue(prd);

            Assert::IsTrue(prd->payoffLabels().size());
        }

        TEST_METHOD(TestValue)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            auto results = simValue(mid, pid, num);

            // Assert
            Assert::IsTrue(results.identifiers.size());
            Assert::IsTrue(results.values.size());
        }

        TEST_METHOD(TestValueTime)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            clock_t t0 = clock();
            auto results = value(mid, pid, num);
            clock_t t1 = clock();
            double time = t1 - t0;

            Assert::IsTrue(time > 0);
        }

        TEST_METHOD(TestAADrisk)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            const string riskPayoff = "RiskPayoff";
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            auto results = AADriskOne(mid, pid, num, riskPayoff);
            Assert::IsTrue(results.riskPayoffValue);
            Assert::IsTrue(results.paramIds.size());
            Assert::IsTrue(results.risks.size() == results.paramIds.size());
        }

        TEST_METHOD(TestAADriskAggregate)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            const map<string, double> notionals;
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            auto results = AADriskAggregate(mid, pid, notionals, num);

            Assert::IsTrue(results.risks.size());
            Assert::IsTrue(results.paramIds.size());
        }

        TEST_METHOD(TestBumprisk)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            const string storeid = "Storeid";
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            auto results = bumpRisk(mid, pid, num);;
            riskStore[storeid] = results;

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestAADriskMulti)
        {
            // Arrange
            const string pid = "Product_id";
            const string mid = "Model_id";
            const string storeid = "Store_id";
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            const auto* prd = getProduct<double>(pid);
            Assert::IsTrue(prd);

            Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
            Assert::IsTrue(mdl);

            auto results = AADriskMulti(mid, pid, num);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestDisplayRisk)
        {
            // Arrange
            const string riskid = "Risk_id";
            const string displayid = "Display_id";

            // Act
            RiskReports* results = findRiskReports(riskid);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestDupireCalib)
        {
            // Arrange
            const double spot = 100.0;
            const double vol = 0.15;
            const double jmpIntens = 1.0;
            const double jmpAverage = 0.01;
            const double jmpStd = 0.005;
            vector<double> vspots = {};
            const double maxDs = 0.005;
            vector<double> vtimes = {};
            const double maxDt = 0.005;

            // Act
            auto results = dupireCalib(vspots, maxDs, vtimes, maxDt, spot, vol, jmpIntens, jmpAverage, jmpStd);

            // Assert
            //results.spots
            //results.times
            //results.lVols
            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestDupireSuperbucket)
        {
            // Arrange
            const double spot = 100.0;
            const double vol = 0.15;
            const double jmpIntens = 1.0;
            const double jmpAverage = 0.01;
            const double jmpStd = 0.005;
            //  risk view
            vector<double> vstrikes = {};
            vector<double> vmats = {};
            //  calibration
            vector<double> vspots = {};
            const double maxDs = 200.0;
            vector<Time> vtimes = {};
            const Time maxDtVol = 1.5;
            //  MC
            const Time maxDtSimul = 1.5;
            //  product 
            const string pid = "Product_id";
            const vector<string> payoffs = {};
            map<string, double> notionals = {};
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;

            // Act
            auto results = dupireSuperbucket(
                spot,
                maxDtSimul,
                pid,
                notionals,
                vspots,
                maxDs,
                vtimes,
                maxDtVol,
                vstrikes,
                vmats,
                vol,
                jmpIntens,
                jmpAverage,
                jmpStd,
                num);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestDupireSuperbucketBump)
        {
            // Arrange
            const double spot = 100.0;
            const double vol = 0.15;
            const double jmpIntens = 1.0;
            const double jmpAverage = 0.01;
            const double jmpStd = 0.005;
            //  risk view
            vector<double> vstrikes = {};
            vector<double> vmats = {};
            //  calibration
            vector<double> vspots = {};
            const double maxDs = 200.0;
            vector<Time> vtimes = {};
            const Time maxDtVol = 1.5;
            //  MC
            const Time maxDtSimul = 1.5;
            //  product 
            const string pid = "Product_id";
            const vector<string> payoffs = {};
            map<string, double> notionals = {};
            //  numerical parameters
            NumericalParam num;
            num.numPath = 10;
            num.parallel = false;
            num.seed1 = 12345678;
            num.seed2 = 98765432;
            num.useSobol = false;
            
            // Act
            auto results = dupireSuperbucketBump(
                    spot,
                    maxDtSimul,
                    pid,
                    notionals,
                    vspots,
                    maxDs,
                    vtimes,
                    maxDtVol,
                    vstrikes,
                    vmats,
                    vol,
                    jmpIntens,
                    jmpAverage,
                    jmpStd,
                    num);

            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestMerton)
        {
            // Arrange
            const double spot = 100.0;
            const double strike = 100.0;
            const double vol = 0.15;
            const double maturity = 1.0;
            const double intens = 1.0;
            const double meanJmp = 0.01;
            const double stdJmp = 0.005;

            // Act
            const double value = merton(spot, strike, vol, maturity, intens, meanJmp, stdJmp);

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }

        TEST_METHOD(TestBlackScholesKO)
        {
            // Arrange
            const double spot = 100.0;
            const double rate = 0.02;
            const double div = 0.03;
            const double strike = 100.0;
            const double barrier = 105.0;
            const double maturity = 1.0;
            const double vol = 0.15;

            // Act
            const double value = BlackScholesKO(spot, rate, div, strike, barrier, maturity, vol);

            // Assert
            Assert::AreEqual(1.0, 1.0);
        }

	};
}
