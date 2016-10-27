#include "test.h"

#include <QFileDialog>

#include "../libOEF/libOEF/liboef.h"

#pragma comment(lib, "libOEF.lib")

test::test(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
}

test::~test()
{

}

void test::on_pushButton_open_clicked()
{
	QString filePath = QFileDialog::getOpenFileName();

	long winId = ui.widget->winId();

	libOEF::instance()->open(winId, filePath, false, "PowerPoint.Show");


}
