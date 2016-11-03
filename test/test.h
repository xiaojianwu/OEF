#ifndef TEST_H
#define TEST_H

#include <QtWidgets/QMainWindow>
#include "ui_test.h"

#include <QResizeEvent>

class test : public QMainWindow
{
	Q_OBJECT

public:
	test(QWidget *parent = 0);
	~test();

private Q_SLOTS:
void on_pushButton_open_clicked();
void on_pushButton_close_clicked();

void on_pushButton_open_2_clicked();
void on_pushButton_close_2_clicked();


protected:
	virtual void resizeEvent(QResizeEvent *event);

private:
	Ui::testClass ui;
};

#endif // TEST_H
