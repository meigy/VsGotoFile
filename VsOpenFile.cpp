#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <atlbase.h>
#include <iostream> // 添加这一行
#include <string> 
#include <locale>
#include <codecvt> // 添加这一行

// import EnvDTE
#pragma warning(disable : 4278)
#pragma warning(disable : 4146)
#import "libid:80cc9f66-e7d8-4ddd-85b6-d9e6cd0e93e2" version("8.0") lcid("0") raw_interfaces_only named_guids
#pragma warning(default : 4146)
#pragma warning(default : 4278)

using namespace std;

static int visual_studio_open_file(char const * filename, unsigned int line)
{
	HRESULT result;
	CLSID clsid;
	result = ::CLSIDFromProgID(L"VisualStudio.DTE", &clsid);
	if (FAILED(result))
		return -1;

	CComPtr<IUnknown> punk;
	result = ::GetActiveObject(clsid, NULL, &punk);
	if (FAILED(result))
		return -2;

	CComPtr<EnvDTE::_DTE> DTE;
	DTE = punk;

	CComPtr<EnvDTE::ItemOperations> item_ops;
	result = DTE->get_ItemOperations(&item_ops);
	if (FAILED(result))
		return -3;

	CComBSTR bstrFileName(filename);
	CComBSTR bstrKind(EnvDTE::vsViewKindTextView);
	CComPtr<EnvDTE::Window> window;
	result = item_ops->OpenFile(bstrFileName, bstrKind, &window);
	if (FAILED(result))
		return -4;

    CComPtr<EnvDTE::Document> doc;
    int maxtry = 3;
    do {        
        result = DTE->get_ActiveDocument(&doc);
        if (SUCCEEDED(result)) break;
        Sleep(1000);
    } while (FAILED(result) && (0 < --maxtry));
	if (FAILED(result))
		return -5;

	CComPtr<IDispatch> selection_dispatch;
	result = doc->get_Selection(&selection_dispatch);
	if (FAILED(result))
		return -6;

	CComPtr<EnvDTE::TextSelection> selection;
	result = selection_dispatch->QueryInterface(&selection);
	if (FAILED(result))
		return -7;

	result = selection->GotoLine(line, TRUE);
	if (FAILED(result))
		return -8;

	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	CoInitialize(NULL);
	//bool result = visual_studio_open_file("C:\\Windows\\win.ini", 2);
	//1 对lpCmdLine进行解析， 参数1：文件路径(支持带空格双引号的路径)， 参数2：行号
	//2 调用方法visual_studio_open_file打开文件跳转到行
	//3 如果出现错误弹出错误原因提示
	//4 过程异常处理，保证正确执行CoUninitialize
	
    // 解析lpCmdLine参数
    string cmdLine(lpCmdLine);
    cmdLine.erase(cmdLine.begin(), std::find_if(cmdLine.begin(), cmdLine.end(), [](int ch) {
        return !std::isspace(ch);
    }));
    cmdLine.erase(std::find_if(cmdLine.rbegin(), cmdLine.rend(), [](int ch) {
        return !std::isspace(ch);
    }).base(), cmdLine.end());

    int lineNum = 0;
    string filePath;

    // 从右向左查找第一个空格
    size_t spacePos = cmdLine.rfind(' ');    
    if (spacePos != string::npos)
    {
        string lineStr = cmdLine.substr(spacePos + 1);
        lineNum = stoi(lineStr);
        filePath = cmdLine.substr(0, spacePos);
        filePath.erase(filePath.begin(), find_if(filePath.begin(), filePath.end(), [](int ch) {
            return !isspace(ch);
        }));
    }
    else
    {
        // 如果没有找到空格，则假设整个lpCmdLine是文件路径
        filePath = cmdLine;
        filePath.erase(filePath.begin(), find_if(filePath.begin(), filePath.end(), [](int ch) {
            return !isspace(ch);
        }));
    }
    if (filePath.size() > 1 && filePath.front() == '"' && filePath.back() == '"')
    {
        filePath = filePath.substr(1, filePath.size() - 2);
    }
    
    // 调用visual_studio_open_file打开文件并跳转到行
    int result = -100;
    try
    {
        result = visual_studio_open_file(filePath.c_str(), lineNum);
    }
    catch (const std::exception& e)
    {
        std::cerr << "Exception occurred: " << e.what() << std::endl;
    }
    catch (...)
    {
        std::cerr << "Unknown exception occurred." << std::endl;
    }

    // 如果出现错误，弹出错误原因提示
    if (0 != result)
    {
        std::wstring errorMsg = L"Failed to open file '";
        errorMsg += std::wstring(filePath.begin(), filePath.end());
        errorMsg += L"' at line ";
        errorMsg += std::to_wstring(lineNum);
        errorMsg += L".";
        errorMsg += L"' code: ";
        errorMsg += std::to_wstring(result);
        MessageBox(NULL, errorMsg.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }

	CoUninitialize();
	return 0;
}