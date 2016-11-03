#include "liboef.h"

#include <QHash>

#define INITGUID // Init the GUIDS for the control...

#include "dsoframer.h"

class libOEFPrivate {


public:
	// Office Embed Info
	struct OEInfo
	{
		long hwndContainer;
		QString filePath;
		bool readOnly;
		QString progID;
		CDsoFramerControl* dso;

		OEInfo()
		{
			dso = nullptr;
		}
	};
	QHash<long, OEInfo> m_hashOE;
};

libOEF* libOEF::_instance = nullptr;
libOEF *libOEF::instance()
{
	if (_instance == nullptr)
	{
		_instance = new libOEF;
	}
	
	return _instance;
}


libOEF::libOEF() 
{
	d_ptr = new libOEFPrivate();

}

int libOEF::open(long hwndContainer, QRect rect,  QString filePath, bool readOnly, QString progID)
{
	libOEFPrivate::OEInfo info;
	info.hwndContainer = hwndContainer;
	info.filePath = filePath;
	info.readOnly = readOnly;
	info.progID = progID;

	CDsoFramerControl* dso = new CDsoFramerControl;
	info.dso = dso;

	d_ptr->m_hashOE[hwndContainer] = info;

	RECT rectDst;
	rectDst.left = rect.left();
	rectDst.top = rect.top();
	rectDst.right = rect.right();
	rectDst.bottom = rect.bottom();
	

	return dso->Open((LPWSTR)filePath.utf16(), readOnly, (LPWSTR)progID.utf16(), (HWND)hwndContainer, rectDst);
}

void libOEF::close(long hwndContainer)
{
	if (d_ptr->m_hashOE.contains(hwndContainer))
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->Close();
			d_ptr->m_hashOE.remove(hwndContainer);
		}
	}
}

void libOEF::active(long hwndContainer)
{
	if (d_ptr->m_hashOE.contains(hwndContainer))
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			info.dso->Activate();
		}
	}
}

void libOEF::resize(long hwndContainer, QRect rect)
{
	if (d_ptr->m_hashOE.contains(hwndContainer))
	{
		libOEFPrivate::OEInfo info = d_ptr->m_hashOE[hwndContainer];
		if (info.dso)
		{
			RECT rectDst;
			rectDst.left = rect.left();
			rectDst.top = rect.top();
			rectDst.right = rect.right();
			rectDst.bottom = rect.bottom();

			info.dso->OnResize(rectDst);
		}
	}
}

void libOEF::release()
{

}
