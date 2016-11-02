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
	QRect rect = ui.widget->rect();

	libOEF::instance()->open(winId, rect, filePath, false, "PowerPoint.Show");


}

void test::on_pushButton_close_clicked()
{
	long winId = ui.widget->winId();
	libOEF::instance()->close(winId);
}

void test::resizeEvent(QResizeEvent *event)
{
	event->accept();

	long winId = ui.widget->winId();
	QRect rect = ui.widget->rect();
	libOEF::instance()->resize(winId, rect);
}

