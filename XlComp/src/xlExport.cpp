
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

//  Excel export wrappers to functions in main.h

#pragma warning(disable:4996)

#include "ThreadPool.h"
#include "main.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <vector>
#include <string>
#include <filesystem>

#include <Python.h>

#include "xlcall.h"
#include "xlframework.h"
#include "xlOper.h"

using namespace std;
namespace fs = std::filesystem;

const string XLCOMP_UTILS_MODULE_NAME = "xlcomp_utils";
const string SABR_EXPORT_MODULE_NAME = "sabr_export";

vector<float> listTupleToVector_Float(PyObject* incoming) {
    vector<float> data;
    if (PyTuple_Check(incoming)) {
        for (Py_ssize_t i = 0; i < PyTuple_Size(incoming); i++) {
            PyObject* value = PyTuple_GetItem(incoming, i);
            data.push_back(PyFloat_AsDouble(value));
        }
    }
    else {
        if (PyList_Check(incoming)) {
            for (Py_ssize_t i = 0; i < PyList_Size(incoming); i++) {
                PyObject* value = PyList_GetItem(incoming, i);
                data.push_back(PyFloat_AsDouble(value));
            }
        }
        else {
            throw logic_error("Passed PyObject pointer was not a list or tuple!");
        }
    }
    return data;
}

vector<string> listTupleToVector_String(PyObject* incoming) {
    vector<string> data;
    if (PyTuple_Check(incoming)) {
        for (Py_ssize_t i = 0; i < PyTuple_Size(incoming); i++) {
            PyObject* pValue = PyTuple_GetItem(incoming, i);
            PyObject* pStrValue = PyObject_Repr(pValue);
            data.push_back(PyUnicode_AsUTF8(pStrValue));
        }
    }
    else {
        if (PyList_Check(incoming)) {
            for (Py_ssize_t i = 0; i < PyList_Size(incoming); i++) {
                PyObject* pValue = PyList_GetItem(incoming, i);
                PyObject* pStrValue = PyObject_Repr(pValue);
                data.push_back(PyUnicode_AsUTF8(pStrValue));
            }
        }
        else {
            data.push_back("Passed PyObject pointer was not a list or tuple!"); //throw logic_error("Passed PyObject pointer was not a list or tuple!");
        }
    }
    return data;
}

int py_load_module_from_string(const string module_name, const string module_code);

int py_initialize_xl_comp()
{
    if (Py_IsInitialized())
        if (Py_FinalizeEx())
            return -1;

    wchar_t* program = Py_DecodeLocale("xlComp", NULL);
    Py_SetProgramName(program);
    Py_Initialize();
    if (!Py_IsInitialized()) 
        return -1;

    PyObject* pName;
    pName = PyUnicode_DecodeFSDefault("xlutils");
    if (!pName) return -1;

    // Utils for running python from the spreadsheet
    const string module_code =
        "import os, sys\n"
        "def display_sys_paths():\n"
        "    return sys.path\n"
        "\n"
        "def add_sys_path(new_sys_path):\n"
        "    if os.path.exists(new_sys_path) and os.path.isdir(new_sys_path) and new_sys_path not in sys.path:\n"
        "        sys.path.append(new_sys_path)\n"
        "\n";
    auto rc = py_load_module_from_string(XLCOMP_UTILS_MODULE_NAME, module_code);
    if (rc) return -1;

    rc = PyRun_SimpleString("import xlcomp_utils");

 /*   PyObject* pModule;
    pModule = PyImport_Import(pName);
    Py_DECREF(pName);
    if (!pModule) return -1;

    pFunc = PyObject_GetAttrString(pModule, argv[2]);   /* pFunc is a new reference * /

    if (pFunc && PyCallable_Check(pFunc)) {
        pArgs = PyTuple_New(argc - 3);
        for (i = 0; i < argc - 3; ++i) {
            pValue = PyLong_FromLong(atoi(argv[i + 3]));
            if (!pValue) {
                Py_DECREF(pArgs);
                Py_DECREF(pModule);
                fprintf(stderr, "Cannot convert argument\n");
                return 1;
            }
            /* pValue reference stolen here: * /
            PyTuple_SetItem(pArgs, i, pValue);
        }
        pValue = PyObject_CallObject(pFunc, pArgs);
        Py_DECREF(pArgs);
        if (pValue != NULL) {
            printf("Result of call: %ld\n", PyLong_AsLong(pValue));
            Py_DECREF(pValue);
        }
        else {
            Py_DECREF(pFunc);
            Py_DECREF(pModule);
            PyErr_Print();
            fprintf(stderr, "Call failed\n");
            return 1;
        }
    }
    else {
        if (PyErr_Occurred())
            PyErr_Print();
        fprintf(stderr, "Cannot find function \"%s\"\n", argv[2]);
    }
//  Py_XDECREF(pFunc);
    Py_DECREF(pModule);*/
    return 0;
}


int py_load_module_from_string(const string module_name, const string module_code)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* builtins = NULL;
    PyObject* pDict = NULL;
    PyObject* pModule = NULL;
    PyObject* pModules = PyImport_GetModuleDict();

    if ((pModule = PyDict_GetItemString(pModules, module_name.c_str())) &&
        PyModule_Check(pModule))
        // Already loaded; don't try to reload
        return 0;

    if (!(pModule = PyModule_New(module_name.c_str())))
        return -1;
    if (PyDict_SetItemString(pModules, module_name.c_str(), pModule) != 0) 
    {
        Py_DECREF(pModule);
        return -1;
    }

    // Set properties on the new module object
    PyModule_AddStringConstant(pModule, "__file__", "");
    pDict = PyModule_GetDict(pModule);      // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    builtins = PyEval_GetBuiltins();        // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    PyDict_SetItemString(pDict, "__builtins__", builtins);

    // Define code in the newly created module
    PyObject* pValue = PyRun_String(module_code.c_str(), Py_file_input, pDict, pDict);
    if (!pValue) 
    {
        Py_DECREF(pModule);
        return -1;
    }

    Py_DECREF(pValue);
    Py_DECREF(pModule);

    return 0;
}


int py_load_module_from_file1(string module_name, fs::path module_path)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* pModules = PyImport_GetModuleDict();
    PyObject* pModule = NULL;

    if ((pModule = PyDict_GetItemString(pModules, module_name.c_str())) &&
        PyModule_Check(pModule))
        // Already loaded; don't try to reload
        return 0;

    if (!(pModule = PyModule_New(module_name.c_str())))
        return -1;

    return -1;
    /*PyObject* module = Py_CompileString(
        // language: Python
        R"EOT(
fake_policy = {
      "maxConnections": 4,
      "policyDir": "/tmp",
      "enableVhostPolicy": True,
      "enableVhostNamePatterns": False,
})EOT",
        "test_module", Py_file_input);
    if (module == nullptr) return TempErr12(xlerrNA);

    PyObject* pModuleObj = PyImport_ExecCodeModule("test_module", module);
    if (pModuleObj == nullptr) return TempErr12(xlerrNA);;
    // better to check with an if, use PyErr_Print() or such to read the error

    PyObject* pAttrObj = PyObject_GetAttrString(pModuleObj, "fake_policy");  /* retuwns a new reference * /
    if(pAttrObj == nullptr) return TempErr12(xlerrNA);

    auto* entity = reinterpret_cast<qd_entity_t*>(pAttrObj);
    REQUIRE(qd_entity_has(entity, "enableVhostNamePatterns"));  // sanity check for the test input
    PyState_AddModule();
    if (1) return TempErr12(xlerrNA);*/
}


int py_load_module_from_file(const string module_name, const string module_path)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* pValue = NULL;
    PyObject* builtins = NULL;
    PyObject* pDict = NULL;
    PyObject* pModule = NULL;
    PyObject* pModules = PyImport_GetModuleDict();

    if ((pModule = PyDict_GetItemString(pModules, module_name.c_str())) &&
        PyModule_Check(pModule))
        // Already loaded; don't try to reload
        return 0;

    if (!(pModule = PyModule_New(module_name.c_str())))
        return -1;

    // Set properties on the new module object
    PyModule_AddStringConstant(pModule, "__file__", "");
    pDict = PyModule_GetDict(pModule);      // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    builtins = PyEval_GetBuiltins();        // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    PyDict_SetItemString(pDict, "__builtins__", builtins);
    // Define code in the newly created module
    const char* MyModuleCode = "print 'Hello world!'";
    pValue = PyRun_String(MyModuleCode, Py_file_input, pDict, pDict);
    if (pValue == NULL) {
        // Handle error
        Py_DECREF(pModule);
        return -1;
    }

    Py_DECREF(pValue);
    Py_DECREF(pModule);

    return 0;
}

const vector<string> py_sys_paths()
{
    if (!Py_IsInitialized())
        return vector<string>();

    /* compile our function
    stringstream buf;
    buf << "def display_sys_paths() :" << endl
        << "    import sys" << endl
        << "    return list(sys.path)" << endl;
    PyObject* pCompiledFn = Py_CompileString(buf.str().c_str(), "", Py_file_input);
    if (!pCompiledFn) return vector<string>();

    // create a module
    PyObject* pModule = PyImport_ExecCodeModule(XLCOMP_UTILS_MODULE_NAME.c_str(), pCompiledFn);
    if (!pModule) return vector<string>();

    ...
    */
    PyObject* pRetValue = NULL;
    PyObject* pFunc = NULL;
    PyObject* pModule = NULL;

    PyObject* pModules = PyImport_GetModuleDict();
    if (!(pModule = PyDict_GetItemString(pModules, XLCOMP_UTILS_MODULE_NAME.c_str())) ||
        !PyModule_Check(pModule) ||
        !(pFunc = PyObject_GetAttrString(pModule, "display_sys_paths")) ||   // retuwns a new reference
        !PyCallable_Check(pFunc))
    {
    //  Py_XDECREF(pFunc);
        return vector<string>();
    }

    pRetValue = PyObject_CallNoArgs(pFunc);
    if (!pRetValue)
    {
        Py_DECREF(pFunc);
        return vector<string>();
    }
    const vector<string> sys_paths = listTupleToVector_String(pRetValue);

    Py_DECREF(pRetValue);
    Py_DECREF(pFunc);
    return sys_paths;
}

int py_add_sys_path(const string new_sys_path)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* pVal1 = NULL;
    PyObject* pFunc = NULL;
    PyObject* pModule = NULL;

    PyObject* pModules = PyImport_GetModuleDict();
    if (!(pModule = PyDict_GetItemString(pModules, XLCOMP_UTILS_MODULE_NAME.c_str())) ||
        !PyModule_Check(pModule) ||
        !(pFunc = PyObject_GetAttrString(pModule, "add_sys_path")) ||       // returns new reference
        !PyCallable_Check(pFunc))
    {
    //  Py_XDECREF(pFunc);
        return -1;          // fprintf(stderr, "Cannot find function \"%s\"\n", argv[2]);
    }

    // convert the first command-line argument to an string
    if (!(pVal1 = PyUnicode_FromString(new_sys_path.c_str())))
    {
        Py_DECREF(pFunc);
        return -1;
    }

    PyObject_CallOneArg(pFunc, pVal1);
    Py_DECREF(pVal1);
    Py_DECREF(pFunc);
    return 0;
}

double py_eval_to_double(const string python_scriptlet, const string return_variable)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* evalModule = NULL;
    PyObject* evalDict = NULL;
    PyObject* evalVal = NULL;

    auto rc = PyRun_SimpleString(python_scriptlet.c_str());
    if (rc)
        return -1;

    if (return_variable.empty())
        return 0;

    evalModule = PyImport_AddModule((char*)"__main__");
    evalDict = PyModule_GetDict(evalModule);           // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    evalVal = PyDict_GetItemString(evalDict, return_variable.c_str());
    if (evalVal == NULL)
    {
    //  Py_DECREF(evalModule);
    //  Py_DECREF(evalDict);
        return -2; // Cannot find that variable
    }
    if (!PyFloat_Check(evalVal))
    {
    //  Py_DECREF(evalModule);
    //--Py_DECREF(evalVal);
        return -3;
    }
    double result = PyFloat_AsDouble(evalVal);
//  Py_DECREF(evalModule);
//--Py_DECREF(evalVal);
    return result;
}

const string py_eval_to_string(const string python_scriptlet, const string return_variable)
{
    if (!Py_IsInitialized())
        return string();

    PyObject* evalModule = NULL;
    PyObject* evalDict = NULL;
    PyObject* evalVal = NULL;

    int rc = PyRun_SimpleString(python_scriptlet.c_str());
    if (rc)
        return "Script failed";

    if (return_variable.empty())
        return "OK";

    evalModule = PyImport_AddModule((char*)"__main__");
    evalDict = PyModule_GetDict(evalModule);                // Returns a borrowed reference: no need to Py_DECREF() it once we are done
    evalVal = PyDict_GetItemString(evalDict, return_variable.c_str());
    if (evalVal == NULL)
    {
    //  Py_DECREF(evalModule);
    //  Py_DECREF(evalDict);
        return "Cannot find that variable"; // Cannot find that variable
    }
    if (!PyUnicode_Check(evalVal))
    {
    //  Py_DECREF(evalModule);
    //--Py_DECREF(evalVal);
        return "Result is not a string"; 
    }
    const string result = PyUnicode_AsUTF8(evalVal);
//  Py_DECREF(evalModule);
//--Py_DECREF(evalVal);
    return result;
}

int py_call_int_func(const string function_name, int param)
{
    if (!Py_IsInitialized())
        return -1;

    //
    char* procname = new char[function_name.length() + 1];
    std::strcpy(procname, function_name.c_str());
    LPXLOPER12 xlReturn;

    PyObject* pName, * pModule, * pDict, * pFunc, * pValue = nullptr, * presult = nullptr;
    // Initialize the Python Interpreter

    // Build the name object
    pName = PyUnicode_FromString((char*)"PythonCode");
    // Load the module object
    pModule = PyImport_Import(pName);
    // pDict is a borrowed reference 
    pDict = PyModule_GetDict(pModule);
    // pFunc is also a borrowed reference 
    pFunc = PyDict_GetItemString(pDict, procname);
    if (PyCallable_Check(pFunc))
    {
        pValue = Py_BuildValue("(i)", (int)(param + EPS));
        PyErr_Print();
        presult = PyObject_CallObject(pFunc, pValue);
        PyErr_Print();
    }
    else
    {
        PyErr_Print();
    }
    //printf("Result is %d\n", _PyLong_AsInt(presult));
    Py_DECREF(pValue);
    // Clean up
    Py_DECREF(pModule);
    Py_DECREF(pName);
    // Finish the Python Interpreter


    // clean 
    delete[] procname;

    return _PyLong_AsInt(presult);
}

double sabr_2002_lognormal_volatility(
    double tau,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu,
    double strike)
{
    if (!Py_IsInitialized())
        return -1;

    PyObject* pRetValue = NULL;
    PyObject* pKwargs = NULL;
    PyObject* pArgs = NULL;
    PyObject* pFunc = NULL;
    PyObject* pModule = NULL;

    PyObject* pModules = PyImport_GetModuleDict();
    if (!(pModule = PyDict_GetItemString(pModules, SABR_EXPORT_MODULE_NAME.c_str())) ||
        !PyModule_Check(pModule) ||
        !(pFunc = PyObject_GetAttrString(pModule, "sabr_2002_lognormal_volatility")) ||     // returns new reference
        !PyCallable_Check(pFunc))
    {
    //  Py_XDECREF(pFunc);
        return -1;          // fprintf(stderr, "Cannot find function \"%s\"\n", argv[2]);
    }

    // create a new tuple with 1 elements
    pArgs = PyTuple_New(7);
    pKwargs = PyDict_New();
    PyTuple_SetItem(pArgs, 0, PyFloat_FromDouble(tau));
    PyTuple_SetItem(pArgs, 1, PyFloat_FromDouble(forward));
    PyTuple_SetItem(pArgs, 2, PyFloat_FromDouble(alpha));
    PyTuple_SetItem(pArgs, 3, PyFloat_FromDouble(beta));
    PyTuple_SetItem(pArgs, 4, PyFloat_FromDouble(rho));
    PyTuple_SetItem(pArgs, 5, PyFloat_FromDouble(nu));
    PyTuple_SetItem(pArgs, 6, PyFloat_FromDouble(strike));

    pRetValue = PyObject_Call(pFunc, pArgs, pKwargs);
    if (!pRetValue)
    {
        Py_DECREF(pKwargs);
        Py_DECREF(pArgs);
        Py_DECREF(pFunc);
        return -1;
    }
    double result = PyFloat_AsDouble(pRetValue);

    Py_DECREF(pRetValue);
    Py_DECREF(pKwargs);
    Py_DECREF(pArgs);
    Py_DECREF(pFunc);

    return result;
}

//  Helpers
NumericalParam xl2num(
    const double              useSobol,
    const double              seed1,
    const double              seed2,
    const double              numPath,
    const double              parallel)
{
    NumericalParam num;

    num.numPath = static_cast<int>(numPath + EPS);
    num.parallel = parallel > EPS;
	if (seed1 >= 1)
	{
		num.seed1 = static_cast<int>(seed1 + EPS);
	}
	else
	{
		num.seed1 = 1234;
	}

	if (seed2 >= 1)
	{
		num.seed2 = static_cast<int>(seed2 + EPS);
	}
	else
	{
		num.seed2 = num.seed1 + 1;
	}

	num.useSobol = useSobol > EPS;

    return num;
}

//	Wrappers

//  change number of threads in the pool
extern "C" __declspec(dllexport)
double xRestartThreadPool(
    double              xNthread)
{
    const int numThread = int(xNthread + EPS);
    // Todo: restartThreadPool(numThread);
    ThreadPool::getInstance()->stop();
    ThreadPool::getInstance()->start(numThread);

    return numThread;
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutBlackScholes(
    double              spot,
    double              vol,
    double              qSpot,
    double              rate,
    double              div,
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    //  Call and return
    putBlackScholes(spot, vol, qSpot > 0, rate, div, id);

    return TempStr12(id);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutDupire(
    //  model parameters
    double              spot,
    FP12 * spots,
    FP12 * times,
    FP12 * vols,
    double              maxDt,
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (maxDt <= 0.0 || id.empty()) 
        return TempErr12(xlerrNA);

    //  Unpack

    if (spots->rows * spots->columns * times->rows * times->columns != vols->rows * vols->columns)
    {
        return TempErr12(xlerrNA);
    }

    vector<double> vspots = to_vector(spots);
    vector<double> vtimes = to_vector(times);
    matrix<double> vvols = to_matrix(vols);

    //  Call and return
    putDupire(spot, vspots, vtimes, vvols, maxDt, id);

    return TempStr12(id);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutEuropean(
    double              strike,
    double              exerciseDate,
    double              settlementDate,
    LPXLOPER12         xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    if (settlementDate <= 0) settlementDate = exerciseDate;

    //  Call and return
    putEuropean(strike, exerciseDate, settlementDate, id);

    return TempStr12(id);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutBarrier(
    double              strike,
    double              barrier,
    double              maturity,
    double              monitorFreq,
    double              smoothing,
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    //  Call and return
    putBarrier(strike, barrier, maturity, monitorFreq, smoothing, id);

    return TempStr12(id);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutContingent(
    double              coupon,
    double              maturity,
    double              payFreq,
    double              smoothing,
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    //  Call and return
    putContingent(coupon, maturity, payFreq, smoothing, id);

    return TempStr12(id);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPutEuropeans(
    FP12*               maturities,
    FP12*               strikes,
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);

    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    //  Make sure we have options
    if (!strikes->rows || !strikes->columns) 
        return TempErr12(xlerrNA);

    //  Maturities and strikes, removing blanks
    vector<double> vmats;
    vector<double> vstrikes;
    {
        size_t rows = maturities->rows;
        size_t cols = maturities->columns;
        
        if (strikes->rows * strikes->columns != rows * cols) 
            return TempErr12(xlerrNA);
        
        double* mats= maturities->array;
        double* strs = strikes->array;

        for (size_t i = 0; i < cols * rows; ++i)
        {
            if (mats[i] > EPS && strs[i] > EPS)
            {
                vmats.push_back(mats[i]);
                vstrikes.push_back(strs[i]);
            }
        }
    }

    //  Make sure we still have options
    if (vmats.empty()) 
        return TempErr12(xlerrNA);

    //  Call and return
    putEuropeans(vmats, vstrikes, id);

    return TempStr12(id);
}

//  Access payoff identifiers and parameters

extern "C" __declspec(dllexport)
LPXLOPER12 xParameters(
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);
    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    // Todo:  Model<double>* mdl = getModel(id)
    Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(id));
    //  Make sure we have a model
    if (!mdl) 
        return TempErr12(xlerrNA);

    const auto& paramLabels = mdl->parameterLabels();
    const auto& params = mdl->parameters();
    vector<double> paramsCopy(params.size());
    transform(params.begin(), params.end(), paramsCopy.begin(), [](const double* p) { return *p; });

    return from_labelsAndNumbers(paramLabels, paramsCopy);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xPayoffIds(
    LPXLOPER12          xid)
{
    FreeAllTempMemory();

    const string id = getString(xid);
    //  Make sure we have an id
    if (id.empty()) 
        return TempErr12(xlerrNA);

    const auto* prd = getProduct<double>(id);
    //  Make sure we have a product
    if (!prd) 
        return TempErr12(xlerrNA);

    return from_strVector(prd->payoffLabels());
}

extern "C" __declspec(dllexport)
LPXLOPER12 xValue(
    LPXLOPER12          modelid,
    LPXLOPER12          productid,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel)
{
    FreeAllTempMemory();

    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    const auto* prd = getProduct<double>(pid);
    //  Make sure we have a product
    if (!prd) return TempErr12(xlerrNA);

    const string mid = getString(modelid);
    //  Make sure we have an id
    if (mid.empty()) return TempErr12(xlerrNA);

    Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
    //  Make sure we have a model
    if (!mdl) return TempErr12(xlerrNA);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    //  Call and return;
    try 
    {
        auto results = value(mid, pid, num);
        return from_labelsAndNumbers(results.identifiers, results.values);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

extern "C" __declspec(dllexport)
LPXLOPER12 xValueTime(
	LPXLOPER12          modelid,
	LPXLOPER12          productid,
	//  numerical parameters
	double              useSobol,
	double              seed1,
	double              seed2,
	double              numPath,
	double              parallel)
{
	FreeAllTempMemory();

	const string pid = getString(productid);
	//  Make sure we have an id
	if (pid.empty()) return TempErr12(xlerrNA);

	const auto* prd = getProduct<double>(pid);
	//  Make sure we have a product
	if (!prd) return TempErr12(xlerrNA);

	const string mid = getString(modelid);
	//  Make sure we have an id
	if (mid.empty()) return TempErr12(xlerrNA);

	Model<double>* mdl = const_cast<Model<double>*>(getModel<double>(mid));
	//  Make sure we have a model
	if (!mdl) return TempErr12(xlerrNA);

	//  Numerical params
	const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
	//  Make sure we have a numPath
	if (!num.numPath) return TempErr12(xlerrNA);

	//  Call and return;
	try
	{
		clock_t t0 = clock();
		auto results = value(mid, pid, num);
		clock_t t1 = clock();
		LPXLOPER12 oper = TempXLOPER12();
		resize(oper, results.values.size() + 1, 1);
		for (int i = 0; i < results.values.size(); ++i) setNum(oper, results.values[i], i, 0);
		setNum(oper, t1 - t0, results.values.size(), 0);

		return oper;
	}
	catch (const exception&)
	{
		return TempErr12(xlerrNA);
	}
}

extern "C" __declspec(dllexport)
LPXLOPER12 xAADrisk(
    LPXLOPER12          modelid,
    LPXLOPER12          productid,
    LPXLOPER12          xRiskPayoff,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel)
{
    FreeAllTempMemory();

    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    const string mid = getString(modelid);
    //  Make sure we have an id
    if (mid.empty()) return TempErr12(xlerrNA);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    //  Risk payoff
    const string riskPayoff = getString(xRiskPayoff);

    try
    {
        auto results = AADriskOne(mid, pid, num, riskPayoff);
		const size_t n = results.risks.size(), N = n + 1;

        LPXLOPER12 oper = TempXLOPER12();
        resize(oper, N, 2);

        setString(oper, "value", 0, 0);
        setNum(oper, results.riskPayoffValue, 0, 1);

        for (size_t i = 0; i < n; ++i)
        {
            setString(oper, results.paramIds[i], i + 1, 0);
            setNum(oper, results.risks[i], i + 1, 1);
        }

        return oper;
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

extern "C" __declspec(dllexport)
LPXLOPER12 xAADriskAggregate(
    LPXLOPER12          modelid,
    LPXLOPER12          productid,
    LPXLOPER12          xPayoffs,
    FP12*               xNotionals,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel)
{
    FreeAllTempMemory();

    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    const string mid = getString(modelid);
    //  Make sure we have an id
    if (mid.empty()) return TempErr12(xlerrNA);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    //  Payoffs and notionals, removing blanks
    map<string, double> notionals;
    {
        size_t rows = getRows(xPayoffs);
        size_t cols = getCols(xPayoffs);

        if (rows * cols == 0 ||
            xNotionals->rows * xNotionals->columns != rows * cols) 
                return TempErr12(xlerrNA);

        size_t idx = 0;
        for (size_t i = 0; i < rows; ++i) for (size_t j = 0; j < cols; ++j)
        {
            string payoff = getString(xPayoffs, i, j);
            double notional = xNotionals->array[idx++];
            if (payoff != "" && fabs(notional) > EPS)
            {
                notionals[payoff] = notional;
            }
        }
    }

    try
    {
        auto results = AADriskAggregate(mid, pid, notionals, num);
        const size_t n = results.risks.size(), N = n + 1;

        LPXLOPER12 oper = TempXLOPER12();
        resize(oper, N, 2);

        setString(oper, "value", 0, 0);
        setNum(oper, results.riskPayoffValue, 0, 1);

        for (size_t i = 0; i < n; ++i)
        {
            setString(oper, results.paramIds[i], i + 1, 0);
            setNum(oper, results.risks[i], i + 1, 1);
        }

        return oper;
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

unordered_map<string, RiskReports> riskStore;

extern "C" __declspec(dllexport)
LPXLOPER12 xBumprisk(
    LPXLOPER12          modelid,
    LPXLOPER12          productid,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel,
    //  display now or put in memory?
    double              displayNow,
    LPXLOPER12          storeid)
{
    FreeAllTempMemory();

    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    const string mid = getString(modelid);
    //  Make sure we have an id
    if (mid.empty()) return TempErr12(xlerrNA);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    try
    {
        auto results = bumpRisk(mid, pid, num);
        if (displayNow > 0.5)
        {
            return from_labelledMatrix(results.params, results.payoffs, results.risks, "value", results.values);
        }
        else
        {
            const string riskId = getString(storeid);
            if (riskId == "") return TempErr12(xlerrNA);
            riskStore[riskId] = results;
            return storeid;
        }
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

extern "C" __declspec(dllexport)
LPXLOPER12 xAADriskMulti(
    LPXLOPER12          modelid,
    LPXLOPER12          productid,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel,
    //  display now or put in memory?
    double              displayNow,
    LPXLOPER12          storeid)
{
    FreeAllTempMemory();

    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    const string mid = getString(modelid);
    //  Make sure we have an id
    if (mid.empty()) return TempErr12(xlerrNA);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    try
    {
        auto results = AADriskMulti(mid, pid, num);
        if (displayNow > 0.5)
        {
            return from_labelledMatrix(results.params, results.payoffs, results.risks, "value", results.values);
        }
        else
        {
            const string riskId = getString(storeid);
            if (riskId == "") return TempErr12(xlerrNA);
            riskStore[riskId] = results;
            return storeid;
        }
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

extern "C" __declspec(dllexport)
LPXLOPER12 xDisplayRisk(
    LPXLOPER12          riskid,
    LPXLOPER12          displayid)
{
    FreeAllTempMemory();

    RiskReports* results;
    const string riskId = getString(riskid);
    const auto& it = riskStore.find(riskId);
    if (it == riskStore.end()) return TempErr12(xlerrNA);
    else results = &it->second;

    const vector<string> riskIds = to_strVector(displayid);
    
    vector<size_t> riskCols;
    
    for (const auto& id : riskIds)
    {
        auto it2 = find(results->payoffs.begin(), results->payoffs.end(), id);
        if (it2 == results->payoffs.end()) return TempErr12(xlerrNA);
        riskCols.push_back(distance(results->payoffs.begin(), it2));
    }
    if (riskCols.empty()) return TempErr12(xlerrNA);

    const size_t nParam = results->risks.rows();
    const size_t nPayoffs = riskCols.size();

    vector<double> showVals(nPayoffs);
    matrix<double> showRisks(nParam, nPayoffs);
    for (size_t j = 0; j < nPayoffs; ++j)
    {
        for (size_t i = 0; i < nParam; ++i)
        {
            showRisks[i][j] = results->risks[i][riskCols[j]];
        }
        showVals[j] = results->values[riskCols[j]];
    }

    return from_labelledMatrix(results->params, riskIds, showRisks, "value", showVals);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xDupireCalib(
    //  Merton market parameters
    const double spot,
    const double vol,
    const double jmpIntens,
    const double jmpAverage,
    const double jmpStd,
    //  Discretization
    FP12* spots,
    const double maxDs,
    FP12* times,
    const double maxDt)
{
    FreeAllTempMemory();

    //  Make sure the last input is given
    if (maxDs == 0 || maxDt == 0) return 0;

    //  Unpack
    vector<double> vspots;
    {
        size_t rows = spots->rows;
        size_t cols = spots->columns;
        double* numbers = spots->array;

        vspots.resize(rows*cols);
        copy(numbers, numbers + rows*cols, vspots.begin());
    }

    vector<double> vtimes;
    {
        size_t rows = times->rows;
        size_t cols = times->columns;
        double* numbers = times->array;

        vtimes.resize(rows*cols);
        copy(numbers, numbers + rows*cols, vtimes.begin());
    }

    auto results = dupireCalib(vspots, maxDs, vtimes, maxDt,
        spot, vol, jmpIntens, jmpAverage, jmpStd);

    //  Return
    return from_labelledMatrix(results.spots, results.times, results.lVols);
}

extern "C" __declspec(dllexport)
LPXLOPER12 xDupireSuperbucket(
    //  Merton market parameters
    const double spot,
    const double vol,
    const double jmpIntens,
    const double jmpAverage,
    const double jmpStd,
    //  risk view
    FP12* strikes,
    FP12* mats,
    //  calibration
    FP12* spots,
    double maxDs,
    FP12* times,
    double maxDtVol,
    //  MC
    double              maxDtSimul,
    //  product 
    LPXLOPER12          productid,
    LPXLOPER12          xPayoffs,
    FP12*               xNotionals,
    //  numerical parameters
    double              useSobol,
    double              seed1,
    double              seed2,
    double              numPath,
    double              parallel,
    //  bump or AAD?
    double              bump)

{
    FreeAllTempMemory();

    //  Make sure the last input is given
    if (numPath <= 0) return TempErr12(xlerrNA);

    //  Make sure we have a product
    const string pid = getString(productid);
    //  Make sure we have an id
    if (pid.empty()) return TempErr12(xlerrNA);

    //  Unpack
    vector<double> vstrikes = to_vector(strikes);
    vector<double> vmats = to_vector(mats);
    vector<double> vspots = to_vector(spots);
    vector<double> vtimes = to_vector(times);

    //  Numerical params
    const auto num = xl2num(useSobol, seed1, seed2, numPath, parallel);
    //  Make sure we have a numPath
    if (!num.numPath) return TempErr12(xlerrNA);

    //  Payoffs and notionals, removing blanks
    map<string, double> notionals;
    {
        size_t rows = getRows(xPayoffs);
        size_t cols = getCols(xPayoffs);;

        if (rows * cols == 0 ||
            xNotionals->rows * xNotionals->columns != rows * cols)
            return TempErr12(xlerrNA);

        size_t idx = 0;
        for (size_t i = 0; i < rows; ++i) for (size_t j = 0; j < cols; ++j)
        {
            string payoff = getString(xPayoffs, i, j);
            double notional = xNotionals->array[idx++];
            if (payoff != "" && fabs(notional) > EPS)
            {
                notionals[payoff] = notional;
            }
        }
    }

    try
    {
        auto results = bump < EPS
            ? dupireSuperbucket(
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
                num)
            : dupireSuperbucketBump(
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

        //  Build return
        const size_t n = results.vega.rows(), m = results.vega.cols();
        const size_t N = n + 4, M = m + 2;

        LPXLOPER12 oper = TempXLOPER12();
        resize(oper, N, M);

        setString(oper, "value", 0, 0);
        setNum(oper, results.value, 0, 1);
        setString(oper, "delta", 1, 0);
        setNum(oper, results.delta, 1, 1);
        setString(oper, "vega", 2, 0);
        setString(oper, "mats", 2, 1);
        for (size_t i = 0; i < m; ++i) setNum(oper, results.mats[i], 2, 2 + i);
        setString(oper, "strikes", 3, 0);
        for (size_t i = 0; i < n; ++i)
        {
            setNum(oper, results.strikes[i], 4 + i, 0);
            for (size_t j = 0; j < m; ++j) setNum(oper, results.vega[i][j], 4 + i, 2 + j);
        }

        // Return it
        return oper;
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
}

extern "C" __declspec(dllexport)
double xMerton(double spot, double vol, double mat, double strike, double intens, double meanJmp, double stdJmp)
{
    return merton(spot, strike, vol, mat, intens, meanJmp, stdJmp);
}

extern "C" __declspec(dllexport)
double xBarrierBlackScholes(double spot, double rate, double div, double vol, double mat, double strike, double barrier)
{
    return BlackScholesKO(spot, rate, div, strike, barrier, mat, vol);
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyRestart()
{
    // Restart python environment
    LPXLOPER12 xlReturn;

    try
    {
        auto rc = py_initialize_xl_comp();
        if (rc)
            return TempErr12(xlerrNA);

        xlReturn = TempNum12(0);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyDisplayVersion()
{
    // Display python version
    LPXLOPER12 xlReturn;
    vector<string> labels = { "Component:"};
    vector<string> versions = { "Version:"};

    try
    {
        labels.push_back("python");
        versions.push_back(py_eval_to_string("import sys;version=sys.version", "version"));
        xlReturn = from_labelsAndStrings(labels, versions);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyDisplaySysPath()
{
    LPXLOPER12 xlReturn;
    try
    {
        auto sys_paths = py_sys_paths();
        if (sys_paths.size() == 0)
            return TempErr12(xlerrNA);
        xlReturn = from_strVector(sys_paths);

    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyAddSysPath(
    LPXLOPER12 xlAdditionalPath)
{
    LPXLOPER12 xlReturn;
    const string additional_path = getString(xlAdditionalPath);
    if (additional_path.empty())
        return TempErr12(xlerrNA);

    try {
        int rc = py_add_sys_path(additional_path);
        if (rc)
            return TempErr12(xlerrNA);
        xlReturn = TempStr12(additional_path.c_str());
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;

/*
    double r = 0.0;
    const string project_path_str = getString(xlProjectPath);
    fs::path project_path = project_path_str;
    bool is_project_path_defined = !project_path_str.empty() && fs::is_directory(project_path);
    r += is_project_path_defined ? 2.0 : 1.0;
    return TempNum12(r);
    fs::path sabr_module_path = std::string((is_project_path_defined ? project_path_str + "\\" : "") + "sabr_exports.py");
    bool is_sabr_module_defined = !sabr_module_path.empty() && fs::exists(sabr_module_path);
    r += is_sabr_module_defined ? 0.2 : 0.1;
    if (is_project_path_defined)
    {
        r += 1.0;
        auto paths = py_add_sys_path(project_path);
    }

    if (is_sabr_module_defined)
    {
        r += 0.1;
        py_load_module_from_file("sabr_exports", sabr_module_path);
    }*/

    // Display python system paths
/*    const string additional_path = getString(xlAdditionalPath);
/*    const string module_script =
        "import sys, os\n"
        "paths = str(sys.path)\n"

        "def add_sys_path(new_sys_path):\n"
        "    if os.path.exists(new_sys_path) and os.path.isdir(new_sys_path) and new_sys_path not in sys.path:\n"
        "        sys.path.append(new_sys_path)\n"
        "    global paths\n"
        "    paths = str(sys.path)\n"

        + additional_path.empty() ? "new_sys_path = 'none'" : ("new_sys_path = '" + new_path + "'\nadd_sys_path(new_sys_path)\n");
    PyObject* evalModule = NULL;
    PyObject* evalDict = NULL;
    PyObject* evalVal = NULL;
    vector<string> syspaths;
    LPXLOPER12 xlReturn;

    evalModule = PyImport_AddModule((char*)"__main__");
    try
    {
        if (PyRun_SimpleString(module_script.c_str()))
            // Script failed
            return TempStr12("Script failed");
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    evalDict = PyModule_GetDict(evalModule);
    try {
        evalVal = PyDict_GetItemString(evalDict, "paths");
        if (evalVal == NULL)
        {
            // Cannot find that variable
            return TempStr12("Cannot find that variable");
        }
    }
    catch (const exception&)
    {
        return TempStr12("cannot PyDict_GetItemString");
    }

    try {
        /*
        * PyString_AsString returns char* repr of PyObject, which should
        * not be modified in any way...this should probably be copied for
        * safety
        * /
        //const char* my_result = PyUnicode_AsUTF8(evalVal)
        syspaths.push_back(PyUnicode_AsUTF8(evalVal));
    }
    catch (const exception&)
    {
        return  TempStr12("PyUnicode_AsUTF8 failed");
    }

    return from_strVector(syspaths);*/
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyImportModule(
    LPXLOPER12 xlModuleName)
{
    // load a module from file or create from string
    LPXLOPER12 xlReturn;
    const string module_name = getString(xlModuleName);

    if (module_name.empty())
        return TempErr12(xlerrNA);

    try
    {
        PyObject* pModule = NULL;
        PyObject* pName = PyUnicode_DecodeFSDefault(module_name.c_str());
        /* Error checking of pName left out */
        if (!(pModule = PyImport_Import(pName)) || !PyModule_Check(pModule))
        {
            Py_DECREF(pName);
            return TempErr12(xlerrNA);
        }
        Py_DECREF(pName);
        Py_DECREF(pModule);

        xlReturn = TempStr12(module_name.c_str());
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyLoadModuleFromString(
    LPXLOPER12 xlModuleName,
    LPXLOPER12 xlModuleString)
{
    // load a module from file or create from string
    LPXLOPER12 xlReturn;
    const string module_name = getString(xlModuleName);
    const string module_string = getString(xlModuleString);

    if (module_name.empty() || module_string.empty())
        return TempErr12(xlerrNA);

    try
    {
        auto rc = py_load_module_from_string(module_name, module_string);
        if (rc) return TempErr12(xlerrNA);

        xlReturn = TempStr12(module_name.c_str());
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyLoadModuleFromPath(
    LPXLOPER12 xlModuleName,
    LPXLOPER12 xlModulePath)
{
    // load a module from file or create from string
    LPXLOPER12 xlReturn;
    const string module_name = getString(xlModuleName);
    const string module_path = getString(xlModulePath);

    if (module_name.empty() || module_path.empty())
        return TempErr12(xlerrNA);

    try
    {
        if (module_path.empty())
        {

        }
        else
        {
            auto rc = py_load_module_from_file(module_name, module_path);
            if (rc) TempErr12(xlerrNA);
        }

        xlReturn = TempStr12(module_name.c_str());
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyRunSimpleString(
    LPXLOPER12 xlPyScript,
    LPXLOPER12 xlReturnVariable,
    double xlIsReturnDouble = false)
{
    // like : "from time import time,ctime\nprint('Today is',ctime(time()))\n"
    LPXLOPER12 xlReturn;
    const string python_scriptlet = getString(xlPyScript);
    auto return_variable = getString(xlReturnVariable);
    bool is_double_value = xlIsReturnDouble > 0.5;
    if (python_scriptlet.empty())
        return TempErr12(xlerrNA);

    try {
        if (is_double_value)
        {
            auto result = py_eval_to_double(python_scriptlet, return_variable);
            xlReturn = TempNum12(result);
        }
        else
        {
            auto result = py_eval_to_string(python_scriptlet, return_variable);
            xlReturn = TempStr12(result.c_str());
        }
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyCallIntFunc(
    LPXLOPER12 xlFunctionName,
    double param)
{
    LPXLOPER12 xlReturn;
    const string function_name = getString(xlFunctionName);
    if (function_name.empty())
        return TempErr12(xlerrNA);

    try {
        int result = py_call_int_func(function_name, (int)param);
        xlReturn = TempInt12(result);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pyFitSabrToFxMarket(
    double time, 
    double spot, 
    double rd, 
    double rf, 
    double atm_vol, 
    double ms_25d, 
    double rr_25d, 
    double ms_10d, 
    double rr_10d)
{
    LPXLOPER12 xlReturn;
    double forward;
    double alpha;
    double beta;
    double rho;
    double nu;
    double err;
    vector<double> sabr_params(7);

    try {
        
        xlReturn = from_dblVector(sabr_params);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }
    return xlReturn;
}


extern "C" __declspec(dllexport)
LPXLOPER12 pySabrLnVol(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu,
    double strike)
{
    LPXLOPER12 xlReturn;

    try
    {
        xlReturn = TempNum12(sabr_2002_lognormal_volatility(time, forward, alpha, beta, rho, nu, strike));
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pySabrAtmfLnVol(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu)
{
    LPXLOPER12 xlReturn;
    double strike = 1.123;
    double vol = 0.25;

    try
    {
        xlReturn = TempNum12(vol);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}


extern "C" __declspec(dllexport)
LPXLOPER12 pySabrDnsLnVol(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu)
{
    LPXLOPER12 xlReturn;
    double strike = 1.123;
    double vol = 0.25;

    try
    {
        xlReturn = TempNum12(vol);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pySabrDnsStrike(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu)
{
    LPXLOPER12 xlReturn;
    double strike = 1.123;

    try
    {
        xlReturn = TempNum12(strike);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

extern "C" __declspec(dllexport)
LPXLOPER12 pySabrBsDeltaAsStrike(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu,
    double bs_delta,
    double forward_delta = true)
{
    LPXLOPER12 xlReturn;
    double vol = 0.25 + forward_delta;
    bool isForwardDelta = forward_delta > 0.5;

    try
    {
        xlReturn = TempNum12(vol);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

//=============================================================================
// C++ functions
//=============================================================================

extern "C" __declspec(dllexport)
LPXLOPER12 sabrLnVol(
    double time,
    double forward,
    double alpha,
    double beta,
    double rho,
    double nu,
    double strike)
{
    LPXLOPER12 xlReturn;
    double vol = 0.25;

    try 
    {
        xlReturn = TempNum12(vol);
    }
    catch (const exception&)
    {
        return TempErr12(xlerrNA);
    }

    return xlReturn;
}

//	Registers

extern "C" __declspec(dllexport) int xlAutoOpen(void)
{
	XLOPER12 xDLL;

	Excel12f(xlGetName, &xDLL, 0);

    //==========================================================================
    // Register PyQfl python functions
    //==========================================================================

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyRestart"),
        (LPXLOPER12)TempStr12(L"B"),
        (LPXLOPER12)TempStr12(L"pyRestart"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Restart python environment"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyDisplayVersion"),
        (LPXLOPER12)TempStr12(L"Q"),
        (LPXLOPER12)TempStr12(L"pyDisplayVersion"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Display python version"),
        (LPXLOPER12)TempStr12(L""));
    
    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyDisplaySysPath"),
        (LPXLOPER12)TempStr12(L"Q"),
        (LPXLOPER12)TempStr12(L"pyDisplaySysPath"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Display python system paths"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyAddSysPath"),
        (LPXLOPER12)TempStr12(L"QQ"),
        (LPXLOPER12)TempStr12(L"pyAddSysPath"),
        (LPXLOPER12)TempStr12(L"AdditonalPath"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Add path to sys.path"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyImportModule"),
        (LPXLOPER12)TempStr12(L"QQ"),
        (LPXLOPER12)TempStr12(L"pyImportModule"),
        (LPXLOPER12)TempStr12(L"ModuleName"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Import module"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyLoadModuleFromString"),
        (LPXLOPER12)TempStr12(L"QQQ"),
        (LPXLOPER12)TempStr12(L"pyLoadModuleFromString"),
        (LPXLOPER12)TempStr12(L"ModuleName, ModuleString"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Create a module from a string"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyLoadModuleFromPath"),
        (LPXLOPER12)TempStr12(L"QQQ"),
        (LPXLOPER12)TempStr12(L"pyLoadModuleFromPath"),
        (LPXLOPER12)TempStr12(L"ModuleName, ModulePath"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Load module from py file"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyRunSimpleString"),
        (LPXLOPER12)TempStr12(L"QQQB"),
        (LPXLOPER12)TempStr12(L"pyRunSimpleString"),
        (LPXLOPER12)TempStr12(L"PyScript, ReturnVariable, IsReturnDouble"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Run a python script from string"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyCallIntFunc"),
        (LPXLOPER12)TempStr12(L"QQB"),
        (LPXLOPER12)TempStr12(L"pyCallIntFunc"),
        (LPXLOPER12)TempStr12(L"function, param"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Call a python integer function"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pyFitSabrToFxMarket"),
        (LPXLOPER12)TempStr12(L"QBBBBBBBBB"),
        (LPXLOPER12)TempStr12(L"pyFitSabrToFxMarket"),
        (LPXLOPER12)TempStr12(L"time, spot, rd, rf, atm_vol, ms_25d, rr_25d, ms_10d, rr_10d"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Fit SABR To Fx Market"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pySabrLnVol"),
        (LPXLOPER12)TempStr12(L"QBBBBBBB"),
        (LPXLOPER12)TempStr12(L"pySabrLnVol"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu, strike"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Lognormal SABR Vol"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pySabrAtmfLnVol"),
        (LPXLOPER12)TempStr12(L"QBBBBBB"),
        (LPXLOPER12)TempStr12(L"pySabrAtmfLnVol"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"SABR ATM forward ln vol"),
        (LPXLOPER12)TempStr12(L""));

    
    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pySabrDnsLnVol"),
        (LPXLOPER12)TempStr12(L"QBBBBBB"),
        (LPXLOPER12)TempStr12(L"pySabrDnsLnVol"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"SABR delta neutral straddle ln vol"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pySabrDnsStrike"),
        (LPXLOPER12)TempStr12(L"QBBBBBB"),
        (LPXLOPER12)TempStr12(L"pySabrDnsStrike"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Find pySabrDnsStrike"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"pySabrBsDeltaAsStrike"),
        (LPXLOPER12)TempStr12(L"QBBBBBBBB"),
        (LPXLOPER12)TempStr12(L"pySabrBsDeltaAsStrike"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu, bs_delta, forward_delta"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"PyQfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Find SABR strike from bs delta"),
        (LPXLOPER12)TempStr12(L""));

    //==========================================================================
    // Register C++ Qfl functions
    //==========================================================================

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"sabrLnVol"),
        (LPXLOPER12)TempStr12(L"QBBBBBBB"),
        (LPXLOPER12)TempStr12(L"sabrLnVol"),
        (LPXLOPER12)TempStr12(L"time, forward, alpha, beta, rho, nu, strike"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Qfl"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Lognormal SABR Vol"),
        (LPXLOPER12)TempStr12(L""));

    //==========================================================================
    // Register Savine functions
    //==========================================================================

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xRestartThreadPool"),
        (LPXLOPER12)TempStr12(L"BB"),
        (LPXLOPER12)TempStr12(L"xRestartThreadPool"),
        (LPXLOPER12)TempStr12(L"numThreads"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Restarts the thread pool with n threads"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutBlackScholes"),
        (LPXLOPER12)TempStr12(L"QBBBBBQ"),
        (LPXLOPER12)TempStr12(L"xPutBlackScholes"),
        (LPXLOPER12)TempStr12(L"spot, vol, qSpot, rate, div, id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a Black-Scholes in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutDupire"),
        (LPXLOPER12)TempStr12(L"QBK%K%K%BQ"),
        (LPXLOPER12)TempStr12(L"xPutDupire"),
        (LPXLOPER12)TempStr12(L"spot, spots, times, vols, maxDt, id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a Dupire in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutEuropean"),
        (LPXLOPER12)TempStr12(L"QBBBQ"),
        (LPXLOPER12)TempStr12(L"xPutEuropean"),
        (LPXLOPER12)TempStr12(L"strike, exerciseDate, [settlementDate], id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a European call in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutBarrier"),
        (LPXLOPER12)TempStr12(L"QBBBBBQ"),
        (LPXLOPER12)TempStr12(L"xPutBarrier"),
        (LPXLOPER12)TempStr12(L"strike, barrier, maturity, monitoringFreq, [smoothingFactor], id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a European call in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutContingent"),
        (LPXLOPER12)TempStr12(L"QBBBBQ"),
        (LPXLOPER12)TempStr12(L"xPutContingent"),
        (LPXLOPER12)TempStr12(L"coupon, maturity, payFreq, [smoothingFactor], id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a European call in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPutEuropeans"),
        (LPXLOPER12)TempStr12(L"QK%K%Q"),
        (LPXLOPER12)TempStr12(L"xPutEuropeans"),
        (LPXLOPER12)TempStr12(L"maturities, strikes, id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Initializes a collection of European call in memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xPayoffIds"),
        (LPXLOPER12)TempStr12(L"QQ"),
        (LPXLOPER12)TempStr12(L"xPayoffIds"),
        (LPXLOPER12)TempStr12(L"id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"The payoff identifiers in a product"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xParameters"),
        (LPXLOPER12)TempStr12(L"QQ"),
        (LPXLOPER12)TempStr12(L"xParameters"),
        (LPXLOPER12)TempStr12(L"id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"The parameters of a model"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xValue"),
        (LPXLOPER12)TempStr12(L"QQQBBBBB"),
        (LPXLOPER12)TempStr12(L"xValue"),
        (LPXLOPER12)TempStr12(L"modelId, productId, useSobol, [seed1], [seed2], N, [Parallel]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Monte-Carlo valuation"),
        (LPXLOPER12)TempStr12(L""));

	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xValueTime"),
		(LPXLOPER12)TempStr12(L"QQQBBBBB"),
		(LPXLOPER12)TempStr12(L"xValueTime"),
		(LPXLOPER12)TempStr12(L"modelId, productId, useSobol, [seed1], [seed2], N, [Parallel]"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Timed Monte-Carlo valuation"),
		(LPXLOPER12)TempStr12(L""));
	
	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xAADrisk"),
        (LPXLOPER12)TempStr12(L"QQQQBBBBB"),
        (LPXLOPER12)TempStr12(L"xAADrisk"),
        (LPXLOPER12)TempStr12(L"modelId, productId, riskPayoff, useSobol, [seed1], [seed2], N, [Parallel]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"AAD risk report"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xAADriskAggregate"),
        (LPXLOPER12)TempStr12(L"QQQQK%BBBBB"),
        (LPXLOPER12)TempStr12(L"xAADriskAggregate"),
        (LPXLOPER12)TempStr12(L"modelId, productId, payoffs, notionals, useSobol, [seed1], [seed2], N, [Parallel]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"AAD risk report for aggregate book of payoffs"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xBumprisk"),
        (LPXLOPER12)TempStr12(L"QQQBBBBBBQ"),
        (LPXLOPER12)TempStr12(L"xBumprisk"),
        (LPXLOPER12)TempStr12(L"modelId, productId, useSobol, [seed1], [seed2], N, [Parallel], [display?], [storeId]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Bump risk report"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xAADriskMulti"),
        (LPXLOPER12)TempStr12(L"QQQBBBBBBQ"),
        (LPXLOPER12)TempStr12(L"xAADriskMulti"),
        (LPXLOPER12)TempStr12(L"modelId, productId, useSobol, [seed1], [seed2], N, [Parallel], [display?], [storeId]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"AAD risk report for multiple payoffs"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xDisplayRisk"),
        (LPXLOPER12)TempStr12(L"QQQ"),
        (LPXLOPER12)TempStr12(L"xDisplayRisk"),
        (LPXLOPER12)TempStr12(L"reportId, payoffId"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Display risk from stored report"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xDupireSuperbucket"),
        (LPXLOPER12)TempStr12(L"QBBBBBK%K%K%BK%BBQQK%BBBBBB"),
        (LPXLOPER12)TempStr12(L"xDupireSuperbucket"),
        (LPXLOPER12)TempStr12(L"spot, vol, jmpIt, jmpAve, jmpStd, RiskStrikes, riskMats, volSpots, maxDs, volTimes, maxDtVol, maxDtSimul, productId, payoffs, notionals, sobol, s1, s2, numPth, parallel, [bump?]"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Computes the risk of a product in Dupire"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xDupireCalib"),
        (LPXLOPER12)TempStr12(L"QBBBBBK%BK%B"),
        (LPXLOPER12)TempStr12(L"xDupireCalib"),
        (LPXLOPER12)TempStr12(L"spot, vol, jumpIntensity, jumpAverage, jumpStd, spots, maxds, times, mxdt"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Calibrates Dupire"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xMerton"),
        (LPXLOPER12)TempStr12(L"BBBBBBBB"),
        (LPXLOPER12)TempStr12(L"xMerton"),
        (LPXLOPER12)TempStr12(L"spot, vol, mat, strike, intens, meanJmp, stdJmp"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Merton"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"xBarrierBlackScholes"),
        (LPXLOPER12)TempStr12(L"BBBBBBBB"),
        (LPXLOPER12)TempStr12(L"xBarrierBlackScholes"),
        (LPXLOPER12)TempStr12(L"spot, rate, div, vol, mat, strike, barrier"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"Savine Modern Comp Finance"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Merton"),
        (LPXLOPER12)TempStr12(L""));

	/* Free the XLL filename */
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    /*  Start the thread pool   */
    ThreadPool::getInstance()->start(thread::hardware_concurrency() - 1);

    //Initialize the python instance
    auto rc = py_initialize_xl_comp();

	return 1;
}

extern "C" __declspec(dllexport) int xlAutoClose(void)
{
    /*  Stop the thread pool   */
    ThreadPool::getInstance()->stop();

    // Close instance
    if (Py_IsInitialized())
        Py_Finalize();
 
    return 1;
}