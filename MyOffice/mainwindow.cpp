#include "mainwindow.h"
#include "ui_mainwindow.h"

#include    <QPainter>
#include    <QFileDialog>

void MainWindow::paintEvent(QPaintEvent *event)
{
    QPainter painter(this);
    painter.drawPixmap(0,ui->mainToolBar->height(),this->width(),
                       this->height()-ui->mainToolBar->height(),
                       QPixmap(":/images/images/back1.jpg"));
    event->accept();
}

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->tabWidget->setVisible(false);
    ui->tabWidget->clear();
    ui->tabWidget->setTabsClosable(true);
    // 不能进入文档模式
    ui->tabWidget->setDocumentMode(false);

    this->setCentralWidget(ui->tabWidget);
    //this->setWindowState(Qt::WindowFullScreen);
    //this->setWindowState(Qt::WindowMaximized);
    // 无需绘制背景颜色
    //this->setAutoFillBackground(false);
    this->resize(1902,1108); // word打开时的默认大小
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::do_changeTabTitle(QString title)
{
    // 用到tabWidget记录自身index的特性
    int index=ui->tabWidget->currentIndex();
    ui->tabWidget->setTabText(index,title);
}


// 关闭具有编号的嵌入式显示窗口
void MainWindow::on_tabWidget_tabCloseRequested(int index)
{
    if (index<0)
        return;

    QWidget* aForm=ui->tabWidget->widget(index);
    aForm->close();
}

// 嵌入式文本显示窗口
// 在这个slot里面直接新建TformDoc，头文件中包含tformdoc.h，类中不需要包含TFormDoc型的变量
// 后面思路基本一致
void MainWindow::on_actWidgetInsite_triggered()
{
    // 指定主窗口为父容器
    TFormDoc *formDoc = new TFormDoc(this);
    // 关闭时自动删除——这是一个attribute
    formDoc->setAttribute(Qt::WA_DeleteOnClose);

    int cur= ui->tabWidget->addTab(formDoc,
                                  QString::asprintf("Doc %d",ui->tabWidget->count()));
    ui->tabWidget->setCurrentIndex(cur);
    ui->tabWidget->setVisible(true);
    // 这里需要与TFormDoc里面的一个信号连接
    connect(formDoc, &TFormDoc::titleChanged,this, &MainWindow::do_changeTabTitle);
}

// 嵌入式表格显示窗口
void MainWindow::on_actWindowInsite_triggered()
{
    TFormTable *formTable = new TFormTable(this);
    formTable->setAttribute(Qt::WA_DeleteOnClose); //关闭时自动删除
    int cur=ui->tabWidget->addTab(formTable,
                                  QString::asprintf("Table %d",ui->tabWidget->count()));
    ui->tabWidget->setCurrentIndex(cur);
    ui->tabWidget->setVisible(true);
}

// 独立式表格显示窗口
void MainWindow::on_actWindow_triggered()
{
    TFormTable* formTable = new TFormTable(this);
    // 对话框关闭时自动删除对话框对象,用于不需要读取返回值的对话框
    formTable->setAttribute(Qt::WA_DeleteOnClose);
    formTable->setWindowTitle("MyExcel");
    formTable->setWindowIcon(QIcon(":/images/images/iconExcel.bmp"));
    // 如果没有状态栏，就创建状态栏
    //formTable->statusBar();
    formTable->show();
}

// 独立式文本显示窗口
void MainWindow::on_actWidget_triggered()
{
    TFormDoc *formDoc = new TFormDoc(this);
    formDoc->setAttribute(Qt::WA_DeleteOnClose);
    formDoc->setWindowTitle("MyWord");
    formDoc->setWindowIcon(QIcon(":/images/images/iconWord.bmp"));
    //formDoc->resize(1902,1108);
    // 视为独立窗口
    formDoc->setWindowFlag(Qt::Window, true);

    //    formDoc->setWindowFlag(Qt::CustomizeWindowHint,true);
    //    formDoc->setWindowFlag(Qt::WindowMinMaxButtonsHint,true);
    //    formDoc->setWindowFlag(Qt::WindowCloseButtonHint,true);
    //    formDoc->setWindowFlag(Qt::WindowStaysOnTopHint,true);

    //    formDoc->setWindowState(Qt::WindowMaximized);
    //    formDoc->setWindowOpacity(0.9);
    //    formDoc->setWindowModality(Qt::WindowModal);

    formDoc->show();
}

// 可见性设置
void MainWindow::on_tabWidget_currentChanged(int index)
{
    Q_UNUSED(index);
    bool  en=ui->tabWidget->count()>0;
    // 当没有tab时，设置为不可见
    ui->tabWidget->setVisible(en);
}

