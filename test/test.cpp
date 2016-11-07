#include "test.h"

#include <QFileDialog>

#include "../libOEF/libOEF/liboef.h"

#include <QInputDialog>

#include <QAxObject>

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
	//L"PowerPoint.Show" // MSPowerPoint // L"PowerPoint.ShowMacroEnabled"
	ClibOEF::instance()->open(winId, rectDst, (LPWSTR)filePath.utf16(), false, L"PowerPoint.Show");
}

void test::on_pushButton_close_clicked()
{
	long winId = ui.widget->winId();
	ClibOEF::instance()->close(winId);
}


void test::on_pushButton_play_clicked()
{
	long winId = ui.widget->winId();
	ClibOEF::instance()->play(winId);
}

void test::on_pushButton_next_clicked()
{
	long winId = ui.widget->winId();
	ClibOEF::instance()->next(winId);
}

void test::on_pushButton_prev_clicked()
{
	long winId = ui.widget->winId();
	ClibOEF::instance()->prev(winId);



}


void test::on_pushButton_jump_clicked()
{
	int pageNo = QInputDialog::getInt(this, tr("Jump to"), tr("page:"), 2, 1, 99);
	long winId = ui.widget->winId();

	//ClibOEF::instance()->jump(winId, pageNo);

	IDispatch* iface = nullptr;
	HRESULT hr = ClibOEF::instance()->GetActiveDocument(winId, &(iface));

	QAxObject activeDocument(iface);

	QAxObject *slideWindow = activeDocument.querySubObject("SlideShowWindow");
	if (slideWindow)
	{
		QAxObject *view = slideWindow->querySubObject("View");
		if (view)
		{
			view->querySubObject("GotoSlide(int)", pageNo);
		}
	}
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



	long winId2 = ui.widget_2->winId();
	QRect rect2 = ui.widget_2->rect();

	RECT rectDst2;
	rectDst2.left = rect2.left();
	rectDst2.top = rect2.top();
	rectDst2.right = rect2.right();
	rectDst2.bottom = rect2.bottom();

	ClibOEF::instance()->resize(winId, rectDst);
	ClibOEF::instance()->resize(winId2, rectDst2);
}

