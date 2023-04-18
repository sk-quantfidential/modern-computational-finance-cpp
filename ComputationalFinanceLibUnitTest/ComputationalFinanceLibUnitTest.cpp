#include "pch.h"
#include "CppUnitTest.h"
#include "..\CompFinanceLib\src\store.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace ComputationalFinanceLibUnitTest
{
	TEST_CLASS(ComputationalFinanceLibUnitTest)
	{
	public:
		
		TEST_METHOD(TestMethod1)
		{
			double strike = 100.0;
			double barrier = 105.0;
			double maturity = 1.0;
			double monitorFreq = 0.25;
			double smoothing = 1.0;
			const string id = "id";

			//  Call and return
			putBarrier(strike, barrier, maturity, monitorFreq, smoothing, id);

			Assert::AreEqual(1.0, 1.0);
		}
	};
}
