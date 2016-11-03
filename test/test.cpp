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


	RECT rectDst;
	rectDst.left = rect.left();
	rectDst.top = rect.top();
	rectDst.right = rect.right();
	rectDst.bottom = rect.bottom();

	ClibOEF::instance()->open(winId, rectDst, (LPWSTR)filePath.utf16(), false, L"PowerPoint.Show");
}

void test::on_pushButton_close_clicked()
{
	long winId = ui.widget->winId();
	ClibOEF::instance()->close(winId);
}


void test::on_pushButton_open_2_clicked()
{
	QString filePath = QFileDialog::getOpenFileName();

	long winId = ui.widget_2->winId();
	QRect rect = ui.widget_2->rect();

	RECT rectDst;
	rectDst.left = rect.left();
	rectDst.top = rect.top();
	rectDst.right = rect.right();
	rectDst.bottom = rect.bottom();

	ClibOEF::instance()->open(winId, rectDst, (LPWSTR)filePath.utf16(), false, L"PowerPoint.Show");


}

void test::on_pushButton_close_2_clicked()
{
	long winId = ui.widget_2->winId();
	ClibOEF::instance()->close(winId);
}


void test::resizeEvent(QResizeEvent *event)
{
	event->accept();

	long winId = ui.widget->winId();
	QRect rect = ui.widget->rect();

	RECT rectDst;
	rectDst.left = rect.left();
	rectDst.top = rect.top();
	rectDst.right = rect.right();
	rectDst.bottom = rect.bottom();

	ClibOEF::instance()->resize(winId, rectDst);
}

