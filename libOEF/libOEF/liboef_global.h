#ifndef LIBOEF_GLOBAL_H
#define LIBOEF_GLOBAL_H

#include <QtCore/qglobal.h>

#ifdef LIBOEF_LIB
# define LIBOEF_EXPORT Q_DECL_EXPORT
#else
# define LIBOEF_EXPORT Q_DECL_IMPORT
#endif

#endif // LIBOEF_GLOBAL_H
