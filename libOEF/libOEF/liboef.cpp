// libOEF.cpp : ���� DLL Ӧ�ó���ĵ���������
//

#include "stdafx.h"
#include "libOEF.h"


// ���ǵ���������һ��ʾ��
LIBOEF_API int nlibOEF=0;

// ���ǵ���������һ��ʾ����
LIBOEF_API int fnlibOEF(void)
{
    return 42;
}

// �����ѵ�����Ĺ��캯����
// �й��ඨ�����Ϣ������� libOEF.h

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

		OEInfo()
		{
			dso = nullptr;
		}
	};
	std::map<long, OEInfo> m_hashOE;
};

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
	libOEFPrivate::OEInfo info;
	info.hwndContainer = hwndContainer;

	CDsoFramerControl* dso = new CDsoFramerControl;
	info.dso = dso;

	d_ptr->m_hashOE[hwndContainer] = info;

	return dso->Open(filePath, readOnly, progID, (HWND)hwndContainer, rect);
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

void ClibOEF::active(long hwndContainer)
{
	if (d_ptr->m_hashOE.find(hwndContainer) != d_ptr->m_hashOE.end())
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->Activate();
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

void ClibOEF::release()
{

}

