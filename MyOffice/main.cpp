#include "mainwindow.h"

#include <QApplication>
#include <QSplashScreen>
#include <QThread>
#include <QProgressBar>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    a.setApplicationName("MyOffice");
    a.setWindowIcon(QIcon(":/images/images/icon.bmp"));

    QPixmap pixmap(":/images/images/splashScreen.jpg");
    QSplashScreen *splash = new QSplashScreen(pixmap);

    // 添加进度条
    QProgressBar *progBar = new QProgressBar(splash);
    progBar->setMaximum(100);
    progBar->setGeometry(10, splash->height() - 30, splash->width() - 20, 20);
    progBar->show();

    splash->show();
    qApp->processEvents(); // qApp是通用指针

    // 模拟加载过程
    for (int i = 0; i <= 100; i++) {
        progBar->setValue(i);
        qApp->processEvents();
        QThread::msleep(15);
    }

    //加载完成后关闭，并显示mainwindow
    splash->finish(nullptr);
    delete splash;

    MainWindow w;
    w.show();
    return a.exec();
}
