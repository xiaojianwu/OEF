/********************************************************************************
** Form generated from reading UI file 'test.ui'
**
** Created by: Qt User Interface Compiler version 5.6.0
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_TEST_H
#define UI_TEST_H

#include <QtCore/QVariant>
#include <QtWidgets/QAction>
#include <QtWidgets/QApplication>
#include <QtWidgets/QButtonGroup>
#include <QtWidgets/QHBoxLayout>
#include <QtWidgets/QHeaderView>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QPushButton>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QVBoxLayout>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_testClass
{
public:
    QWidget *centralWidget;
    QHBoxLayout *horizontalLayout;
    QVBoxLayout *verticalLayout;
    QWidget *widget;
    QPushButton *pushButton_open;
    QPushButton *pushButton_play;
    QPushButton *pushButton_prev;
    QPushButton *pushButton_next;
    QPushButton *pushButton_jump;
    QPushButton *pushButton_close;
    QPushButton *pushButton_quit;
    QVBoxLayout *verticalLayout_2;
    QWidget *widget_2;
    QPushButton *pushButton_open_2;
    QPushButton *pushButton_close_2;
    QMenuBar *menuBar;
    QToolBar *mainToolBar;
    QStatusBar *statusBar;

    void setupUi(QMainWindow *testClass)
    {
        if (testClass->objectName().isEmpty())
            testClass->setObjectName(QStringLiteral("testClass"));
        testClass->resize(1043, 741);
        centralWidget = new QWidget(testClass);
        centralWidget->setObjectName(QStringLiteral("centralWidget"));
        horizontalLayout = new QHBoxLayout(centralWidget);
        horizontalLayout->setSpacing(6);
        horizontalLayout->setContentsMargins(11, 11, 11, 11);
        horizontalLayout->setObjectName(QStringLiteral("horizontalLayout"));
        verticalLayout = new QVBoxLayout();
        verticalLayout->setSpacing(6);
        verticalLayout->setObjectName(QStringLiteral("verticalLayout"));
        widget = new QWidget(centralWidget);
        widget->setObjectName(QStringLiteral("widget"));

        verticalLayout->addWidget(widget);

        pushButton_open = new QPushButton(centralWidget);
        pushButton_open->setObjectName(QStringLiteral("pushButton_open"));

        verticalLayout->addWidget(pushButton_open);

        pushButton_play = new QPushButton(centralWidget);
        pushButton_play->setObjectName(QStringLiteral("pushButton_play"));

        verticalLayout->addWidget(pushButton_play);

        pushButton_prev = new QPushButton(centralWidget);
        pushButton_prev->setObjectName(QStringLiteral("pushButton_prev"));

        verticalLayout->addWidget(pushButton_prev);

        pushButton_next = new QPushButton(centralWidget);
        pushButton_next->setObjectName(QStringLiteral("pushButton_next"));

        verticalLayout->addWidget(pushButton_next);

        pushButton_jump = new QPushButton(centralWidget);
        pushButton_jump->setObjectName(QStringLiteral("pushButton_jump"));

        verticalLayout->addWidget(pushButton_jump);

        pushButton_close = new QPushButton(centralWidget);
        pushButton_close->setObjectName(QStringLiteral("pushButton_close"));

        verticalLayout->addWidget(pushButton_close);

        pushButton_quit = new QPushButton(centralWidget);
        pushButton_quit->setObjectName(QStringLiteral("pushButton_quit"));

        verticalLayout->addWidget(pushButton_quit);


        horizontalLayout->addLayout(verticalLayout);

        verticalLayout_2 = new QVBoxLayout();
        verticalLayout_2->setSpacing(6);
        verticalLayout_2->setObjectName(QStringLiteral("verticalLayout_2"));
        widget_2 = new QWidget(centralWidget);
        widget_2->setObjectName(QStringLiteral("widget_2"));

        verticalLayout_2->addWidget(widget_2);

        pushButton_open_2 = new QPushButton(centralWidget);
        pushButton_open_2->setObjectName(QStringLiteral("pushButton_open_2"));

        verticalLayout_2->addWidget(pushButton_open_2);

        pushButton_close_2 = new QPushButton(centralWidget);
        pushButton_close_2->setObjectName(QStringLiteral("pushButton_close_2"));

        verticalLayout_2->addWidget(pushButton_close_2);


        horizontalLayout->addLayout(verticalLayout_2);

        testClass->setCentralWidget(centralWidget);
        menuBar = new QMenuBar(testClass);
        menuBar->setObjectName(QStringLiteral("menuBar"));
        menuBar->setGeometry(QRect(0, 0, 1043, 23));
        testClass->setMenuBar(menuBar);
        mainToolBar = new QToolBar(testClass);
        mainToolBar->setObjectName(QStringLiteral("mainToolBar"));
        testClass->addToolBar(Qt::TopToolBarArea, mainToolBar);
        statusBar = new QStatusBar(testClass);
        statusBar->setObjectName(QStringLiteral("statusBar"));
        testClass->setStatusBar(statusBar);

        retranslateUi(testClass);

        QMetaObject::connectSlotsByName(testClass);
    } // setupUi

    void retranslateUi(QMainWindow *testClass)
    {
        testClass->setWindowTitle(QApplication::translate("testClass", "test", 0));
        pushButton_open->setText(QApplication::translate("testClass", "\346\211\223\345\274\200", 0));
        pushButton_play->setText(QApplication::translate("testClass", "\346\222\255\346\224\276", 0));
        pushButton_prev->setText(QApplication::translate("testClass", "\344\270\212\344\270\200\351\241\265", 0));
        pushButton_next->setText(QApplication::translate("testClass", "\344\270\213\344\270\200\351\241\265", 0));
        pushButton_jump->setText(QApplication::translate("testClass", "\350\267\263\350\275\254", 0));
        pushButton_close->setText(QApplication::translate("testClass", "\345\205\263\351\227\255", 0));
        pushButton_quit->setText(QApplication::translate("testClass", "\351\200\200\345\207\272", 0));
        pushButton_open_2->setText(QApplication::translate("testClass", "\346\211\223\345\274\200", 0));
        pushButton_close_2->setText(QApplication::translate("testClass", "\345\205\263\351\227\255", 0));
    } // retranslateUi

};

namespace Ui {
    class testClass: public Ui_testClass {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_TEST_H
