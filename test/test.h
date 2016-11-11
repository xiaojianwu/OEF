#ifndef TEST_H
#define TEST_H

#include <QtWidgets/QMainWindow>
#include "ui_test.h"

#include <QResizeEvent>

#include <QAxObject>

class test : public QMainWindow
{
	Q_OBJECT

public:
	test(QWidget *parent = 0);
	~test();

private Q_SLOTS:
void on_pushButton_open_clicked();
void on_pushButton_close_clicked();

void on_pushButton_play_clicked();

void on_pushButton_next_clicked();

void on_pushButton_prev_clicked();

void on_pushButton_jump_clicked();


void on_pushButton_open_2_clicked();
void on_pushButton_close_2_clicked();

void on_pushButton_next_2_clicked();
void on_pushButton_prev_2_clicked();



private: 
	QAxObject* getSlideView(long winId);


protected:
	virtual void resizeEvent(QResizeEvent *event);

private:
	Ui::testClass ui;

	QAxObject *m_pptApp;
};

#endif // TEST_H
