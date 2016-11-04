// 下列 ifdef 块是创建使从 DLL 导出更简单的
// 宏的标准方法。此 DLL 中的所有文件都是用命令行上定义的 LIBOEF_EXPORTS
// 符号编译的。在使用此 DLL 的
// 任何其他项目上不应定义此符号。这样，源文件中包含此文件的任何其他项目都会将
// LIBOEF_API 函数视为是从 DLL 导入的，而此 DLL 则将用此宏定义的
// 符号视为是被导出的。
#ifdef LIBOEF_EXPORTS
#define LIBOEF_API __declspec(dllexport)
#else
#define LIBOEF_API __declspec(dllimport)
#endif

#include <wtypes.h>

class libOEFPrivate;

// 此类是从 libOEF.dll 导出的
class LIBOEF_API ClibOEF {

private:
	ClibOEF(void);

public:
	static ClibOEF* instance();

	void release();


public:
	int open(long hwndContainer, RECT rect, LPWSTR filePath, bool readOnly = true, LPWSTR progID = L"");

	void play(long hwndContainer);

	void next(long hwndContainer);

	void prev(long hwndContainer);


	void jump(long hwndContainer, int pageNo);

	void close(long hwndContainer);

	void active(long hwndContainer);

	void resize(long hwndContainer, RECT rect);


private:
	libOEFPrivate *d_ptr;
	static ClibOEF* _instance;
};

extern LIBOEF_API int nlibOEF;

LIBOEF_API int fnlibOEF(void);
