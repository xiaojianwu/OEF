// libOEF.cpp : 定义 DLL 应用程序的导出函数。
//

#include "stdafx.h"
#include "libOEF.h"


// 这是导出变量的一个示例
LIBOEF_API int nlibOEF=0;

// 这是导出函数的一个示例。
LIBOEF_API int fnlibOEF(void)
{
    return 42;
}

// 这是已导出类的构造函数。
// 有关类定义的信息，请参阅 libOEF.h

#define INITGUID // Init the GUIDS for the control...

#include "dsoframer.h"

#include <map>

class libOEFPrivate {


public:
	// Office Embed Info
	struct OEInfo
	{
		long hwndContainer;
		CDsoFramerControl* dso;

		IDispatch* ifaceActiveWindow;

		OEInfo()
		{
			dso = nullptr;
			ifaceActiveWindow = nullptr;
		}
	};
	std::map<long, OEInfo> m_hashOE;
};

// 获取ppt播放窗口句柄
HWND GetSliderHwnd(HWND hDSOFramerDocWnd)
{

	// locate the host window of the PowerPoint presentation
	HWND hChildClass = FindWindowEx(hDSOFramerDocWnd, NULL, L"childClass", NULL);
	HWND hWndSliderShow = FindWindowEx(hChildClass, NULL, L"childClass", NULL);

	return hWndSliderShow;
}


ClibOEF* ClibOEF::_instance = nullptr;
ClibOEF *ClibOEF::instance()
{
	if (_instance == nullptr)
	{
		_instance = new ClibOEF;
	}

	return _instance;
}


ClibOEF::ClibOEF()
{
	d_ptr = new libOEFPrivate();

}

int ClibOEF::open(long hwndContainer, RECT rect, LPWSTR filePath, bool readOnly, LPWSTR progID)
{
	// 防止在统一窗口句柄上打开多次，先关闭
	close(hwndContainer);

	libOEFPrivate::OEInfo info;
	info.hwndContainer = hwndContainer;

	CDsoFramerControl* dso = new CDsoFramerControl;
	info.dso = dso;

	d_ptr->m_hashOE[hwndContainer] = info;

	return dso->Open(filePath, readOnly, progID, (HWND)hwndContainer, rect);
}

#define COMMAND_PPT_RUN		 (0x100AC)
#define COMMAND_PPT_PAGEDOWN (0x10189)
#define COMMAND_PPT_PAGEUP	 (0x1018a)




void ClibOEF::play(long hwndContainer)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			// start the slideshow
			SendMessage(info.dso->GetActiveWindow(), WM_COMMAND, COMMAND_PPT_RUN, NULL);

			info.dso->ReobtainActiveFrame();
		}
	}
}

// 下一页
void ClibOEF::next(long hwndContainer)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			SendMessage(info.dso->GetActiveWindow(), WM_COMMAND, COMMAND_PPT_PAGEDOWN, NULL);
		}
	}
}

// 上一页
void ClibOEF::prev(long hwndContainer)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			SendMessage(info.dso->GetActiveWindow(), WM_COMMAND, COMMAND_PPT_PAGEUP, NULL);
		}
	}
}

HRESULT ClibOEF::GetActiveDocument(long hwndContainer, IDispatch** ppdisp)
{
	HRESULT hr = S_OK;
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			hr =  info.dso->GetActiveDocument(ppdisp);
		}
	}
	return hr;
}

void ClibOEF::close(long hwndContainer)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->Close();
			d_ptr->m_hashOE.erase(hwndContainer);
		}
	}
}


void ClibOEF::showToolbars(long hwndContainer, bool bShow)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->ShowToolbars(bShow);
		}
	}
}

void ClibOEF::resize(long hwndContainer, RECT rect)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->OnResize(rect);
		}
	}
}

