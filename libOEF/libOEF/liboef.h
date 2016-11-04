// ���� ifdef ���Ǵ���ʹ�� DLL �������򵥵�
// ��ı�׼�������� DLL �е������ļ��������������϶���� LIBOEF_EXPORTS
// ���ű���ġ���ʹ�ô� DLL ��
// �κ�������Ŀ�ϲ�Ӧ����˷��š�������Դ�ļ��а������ļ����κ�������Ŀ���Ὣ
// LIBOEF_API ������Ϊ�Ǵ� DLL ����ģ����� DLL ���ô˺궨���
// ������Ϊ�Ǳ������ġ�
#ifdef LIBOEF_EXPORTS
#define LIBOEF_API __declspec(dllexport)
#else
#define LIBOEF_API __declspec(dllimport)
#endif

#include <wtypes.h>

class libOEFPrivate;

// �����Ǵ� libOEF.dll ������
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
