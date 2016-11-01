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
	};
	QHash<long, OEInfo> m_hashOE;
};

libOEF* libOEF::_instance = nullptr;
libOEF *libOEF::instance()
{
	_instance = new libOEF;
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

	d_ptr->m_hashOE[hwndContainer] = info;

	CDsoFramerControl* dso = new CDsoFramerControl;

	RECT rectDst;
	rectDst.left = rect.left();
	rectDst.top = rect.top();
	rectDst.right = rect.right();
	rectDst.bottom = rect.bottom();
	

	dso->Open((LPWSTR)filePath.utf16(), readOnly, (LPWSTR)progID.utf16(), (HWND)hwndContainer, rectDst);
	//CDsoDocObject::CreateInstance();

	return 0;
}

void libOEF::close(long hwndContainer)
{

}

void libOEF::release()
{

}
