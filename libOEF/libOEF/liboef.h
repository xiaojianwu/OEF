#ifndef LIBOEF_H
#define LIBOEF_H

#include "liboef_global.h"

#include <QObject>
#include <QString>

class libOEFPrivate;

class LIBOEF_EXPORT libOEF : public QObject
{
	Q_OBJECT
	Q_DECLARE_PRIVATE(libOEF)

public:
	static libOEF* instance();

	void release();


public:
	int open(long hwndContainer, QString filePath, bool readOnly = true, QString progID = "");

	void close(long hwndContainer);


private:
	libOEF();

private:
	libOEFPrivate *d_ptr;
};

#endif // LIBOEF_H
