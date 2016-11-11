#include "test.h"

#include <QFileDialog>

#include "../libOEF/libOEF/liboef.h"

#include <QInputDialog>

#include <QAxObject>

#include <QMessageBox>

#pragma comment(lib, "libOEF.lib")

test::test(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);

	m_activeDocumentLeft = nullptr;
	m_activeDocumentRight = nullptr;
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

	int code = ClibOEF::instance()->open(winId, rectDst, (LPWSTR)filePath.utf16(), false, L"PowerPoint.Show");
	if (code != 0)
	{
		QMessageBox::warning(this, "error", QString("code=%1").arg(code));
		return;
	}

	IDispatch* iface = nullptr;
	HRESULT hr = ClibOEF::instance()->GetActiveDocument(winId, &(iface));

	if (FAILED(hr) && iface == nullptr)
	{
		return;
	}

	m_activeDocumentLeft = new QAxObject(iface);

	ClibOEF::instance()->play(winId);

}

void test::on_pushButton_close_clicked()
{

	delete m_activeDocumentLeft;
	m_activeDocumentLeft = nullptr;


	long winId = ui.widget->winId();
	ClibOEF::instance()->close(winId);
}


void test::on_pushButton_play_clicked()
{
	long winId = ui.widget->winId();

	ClibOEF::instance()->play(winId);


	return;

	QAxObject* SlideShowSettings = m_activeDocumentLeft->querySubObject("SlideShowSettings");
	if (SlideShowSettings) {
		SlideShowSettings->setProperty("LoopUntilStopped", QVariant(true));
		SlideShowSettings->setProperty("ShowType", QVariant(1));
		//// 返回或设置指定幻灯片放映的放映类型。 ppShowTypeWindow ppShowTypeSpeaker ppShowTypeKiosk
		////SlideShowSettings->setProperty("ShowType", ppShowTypeSpeaker);
		SlideShowSettings->dynamicCall("Run()");
	}	
}

void test::on_pushButton_next_clicked()
{

	QAxObject *slideWindow = m_activeDocumentLeft->querySubObject("SlideShowWindow");
	if (slideWindow)
	{
		QAxObject *view = slideWindow->querySubObject("View");
		if (view)
		{
			view->dynamicCall("Next()");
		}
	}
}

void test::on_pushButton_prev_clicked()
{
	QAxObject *slideWindow = m_activeDocumentLeft->querySubObject("SlideShowWindow");
	if (slideWindow)
	{
		QAxObject *view = slideWindow->querySubObject("View");
		if (view)
		{
			view->dynamicCall("Previous()");
		}
	}

}


void test::on_pushButton_jump_clicked()
{
	int pageNo = QInputDialog::getInt(this, tr("Jump to"), tr("page:"), 2, 1, 99);

	QAxObject *slideWindow = m_activeDocumentLeft->querySubObject("SlideShowWindow");
	if (slideWindow)
	{
		QAxObject *view = slideWindow->querySubObject("View");
		if (view)
		{
			view->dynamicCall("GotoSlide(int)", pageNo);
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

	// L"Word.Document"
	ClibOEF::instance()->open(winId, rectDst, (LPWSTR)filePath.utf16(), false);

	ClibOEF::instance()->showToolbars(winId, false);

	// Document 接口 
	IDispatch* iface = nullptr;
	HRESULT hr = ClibOEF::instance()->GetActiveDocument(winId, &(iface));

	if (FAILED(hr) && iface == nullptr)
	{
		return;
	}

	m_activeDocumentRight = new QAxObject(iface);

	QAxObject *pActiveWindow = m_activeDocumentRight->querySubObject("ActiveWindow");
	if (!pActiveWindow)
	{
		return;
	}
	QAxObject *pView = pActiveWindow->querySubObject("View");
	if (!pView)
	{
		return;
	}
	QAxObject *pZoom = pView->querySubObject("Zoom");
	if (!pZoom)
	{
		return;
	}
	QVariant pagetFit = pZoom->property("PageFit");
	pZoom->setProperty("PageFit", QVariant(2));
}

void test::on_pushButton_next_2_clicked()
{
	long winId = ui.widget_2->winId();

	QAxObject *activeWindow = m_activeDocumentRight->querySubObject("ActiveWindow");

	if (activeWindow)
	{
		//activeWindow->dynamicCall("PageScroll(int)", 1);
		//activeWindow->dynamicCall("SmallScroll(int)", 48);
		activeWindow->dynamicCall("LargeScroll(int)", 1);
	}
}

void test::on_pushButton_prev_2_clicked()
{
	long winId = ui.widget_2->winId();

	QAxObject *activeWindow = m_activeDocumentRight->querySubObject("ActiveWindow");

	if (activeWindow)
	{
		//activeWindow->dynamicCall("PageScroll(int, int)", 0, 1);
		//activeWindow->dynamicCall("SmallScroll(int, int)", 0, 48);
		activeWindow->dynamicCall("LargeScroll(int, int)", 0, 1);
	}
}

void test::on_pushButton_close_2_clicked()
{
	delete m_activeDocumentRight;
	m_activeDocumentRight = nullptr;

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


	long winId2 = ui.widget_2->winId();
	QRect rect2 = ui.widget_2->rect();

	RECT rectDst2;
	rectDst2.left = rect2.left();
	rectDst2.top = rect2.top();
	rectDst2.right = rect2.right();
	rectDst2.bottom = rect2.bottom();

	
	ClibOEF::instance()->resize(winId2, rectDst2);
}

