#ifndef LIBOEF_H
#define LIBOEF_H

#include "liboef_global.h"

#include <QObject>
#include <QString>
#include <QRect>

class libOEFPrivate;

class libOEF : public QObject
{
	Q_OBJECT
	Q_DECLARE_PRIVATE(libOEF)

public:
	static libOEF* instance();

	void release();


public:
	int open(long hwndContainer, QRect rect, QString filePath, bool readOnly = true, QString progID = "");

	void close(long hwndContainer);

	void resize(long hwndContainer, QRect rect);


private:
	libOEF();

private:
	libOEFPrivate *d_ptr;
	static libOEF* _instance;
};

#endif // LIBOEF_H
