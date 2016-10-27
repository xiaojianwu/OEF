#include "liboef.h"

#include <QHash>

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

libOEF *libOEF::instance()
{
	static libOEF _instance;
	return &_instance;
}


libOEF::libOEF() 
{
	d_ptr = new libOEFPrivate();
}

int libOEF::open(long hwndContainer, QString filePath, bool readOnly, QString progID)
{
	libOEFPrivate::OEInfo info;
	info.hwndContainer = hwndContainer;
	info.filePath = filePath;
	info.readOnly = readOnly;
	info.progID = progID;

	d_ptr->m_hashOE[hwndContainer] = info;

	return 0;
}

void libOEF::close(long hwndContainer)
{

}

void libOEF::release()
{

}
