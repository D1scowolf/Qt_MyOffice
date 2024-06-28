#include "tformdoc.h"
#include "ui_tformdoc.h"

#include    <QToolBar>
#include    <QVBoxLayout>
#include    <QFileDialog>
#include    <QFile>
#include    <QSaveFile>
#include    <QTextStream>
#include    <QFontDialog>
#include    <QFileInfo>
#include    <QMessageBox>
#include    <QInputDialog>
#include    <QStatusBar>

#include    <mainwindow.h>

TFormDoc::TFormDoc(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::TFormDoc)
{
    ui->setupUi(this);

    // 自行设计ToolBar
    QToolBar* locToolBar = new QToolBar("文档",this);
    locToolBar->addAction(ui->actOpen);
    locToolBar->addAction(ui->actSave);
    locToolBar->addAction(ui->actOpenWord);
    locToolBar->addAction(ui->actSaveWord);
    locToolBar->addSeparator();
    locToolBar->addAction(ui->actFont);
    locToolBar->addAction(ui->actCut);
    locToolBar->addAction(ui->actCopy);
    locToolBar->addAction(ui->actPaste);
    locToolBar->addAction(ui->actUndo);
    locToolBar->addAction(ui->actRedo);
    locToolBar->addSeparator();
    locToolBar->addAction(ui->actClose);

    locToolBar->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);

    statusBar = new QStatusBar();

    QVBoxLayout *Layout = new QVBoxLayout();
    Layout->addWidget(locToolBar);
    Layout->addWidget(ui->plainTextEdit);
    Layout->setContentsMargins(2,2,2,2);    //减小边框的宽度
    Layout->setSpacing(2);
    Layout->addWidget(statusBar);

    // 信号与槽的实现
    connect(ui->plainTextEdit, &QPlainTextEdit::textChanged,
            this, &TFormDoc::do_textChanged);

    this->setLayout(Layout);
}

TFormDoc::~TFormDoc()
{
    // QMessageBox::information(this, "消息", "TFormDoc对象被删除和释放");
    delete ui;
}

// 在字数变化的时候显示当前字数
void TFormDoc::do_textChanged()
{
    int charCount = ui->plainTextEdit->toPlainText().length();
    statusBar->showMessage(QString("当前字数：%1").arg(charCount));
}

// 读取文本文件
void TFormDoc::loadFromFile(QString &aFileName)
{
    QFile aFile(aFileName);
    // 只读方式
    if (aFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        // 用文本流读取文件
        QTextStream aStream(&aFile);
        ui->plainTextEdit->clear();
        // 采用一行行读取的方法
        while (!aStream.atEnd())
        {
            QString str=aStream.readLine();
            ui->plainTextEdit->appendPlainText(str);
        }
        aFile.close();

        // 更改tab的标题
        QFileInfo fileInfo(aFileName);
        QString shortName=fileInfo.fileName();
        this->setWindowTitle(shortName);
        emit titleChanged(shortName);
    }
}

// 写入文本文件
bool TFormDoc::saveByIO_Whole(const QString &aFilename)
{
    QFile aFile(aFilename);
    // 写入方式，检查
    if(!aFile.open(QIODevice::WriteOnly | QIODevice::Text))
        return false;
    // 整体读取
    QString str = ui->plainTextEdit->toPlainText();
    // 转换编码方式
    QByteArray strBytes = str.toUtf8();
    aFile.write(strBytes, strBytes.length());
    aFile.close();
    return true;
}

// 打开文件的action
void TFormDoc::on_actOpen_triggered()
{
    //curPath=QCoreApplication::applicationDirPath();
    QString curPath = QDir::currentPath();
    QString dlgTitle = "打开一个文件";
    QString filter = "h 文件(*.h);;C++文件(*.cpp);;文本文件(*.txt);;xml文件(*.xml);;json文件(*.json)";
    // 标准的点击按钮，打开文件管理器的操作
    QString aFileName = QFileDialog::getOpenFileName(this, dlgTitle, curPath, filter);

    if (aFileName.isEmpty())
        return;

    loadFromFile(aFileName);
}

// 写入文件的action
void TFormDoc::on_actSave_triggered()
{
    QString curPath = QDir::currentPath();
    QString dlgTitle = "保存一个文件";
    QString filter = "h 文件(*.h);;C++文件(*.cpp);;文本文件(*.txt);;xml文件(*.xml);;json文件(*.json)";
    QString aFileName = QFileDialog::getSaveFileName(this, dlgTitle, curPath, filter);

    if (aFileName.isEmpty())
        return;

    QFileInfo fileinfo(aFileName);
    QDir::setCurrent(fileinfo.absolutePath());
    saveByIO_Whole(aFileName);
}

// 打开word的action
void TFormDoc::on_actOpenWord_triggered()
{
    // 获取word应用和文件集合
    myword = new QAxObject("Word.Application");
    mydocs = myword->querySubObject("Documents");

    // 错误处理1
    if (mydocs == nullptr)
    {
        QMessageBox::information(this, QString("标题"), QString("获取所有工作文档失败"));
        return;
    }

//    // 设置的默认路径为桌面，在他人电脑上运行时需要修改路径
//    // fileName为桌面上任意docx文件
//    bool isOK;
//    QString fileName = QInputDialog::getText(NULL, "MyWord", "请输入word文档名称", QLineEdit::Normal, "test.docx", &isOK);
//    // 错误处理2
//    if(!isOK)
//    {
//        QMessageBox::information(this, QString("标题"), QString("输入文档不存在"));
//        return;
//    }
//    QString path = QString(DocPath)+fileName;

    // 不需要自定义路径了
    QString curPath = QDir::currentPath();
    QString dlgTitle = "打开一个word文档";
    QString filter = "word 文档(*.docx);;word 文档(*.doc)";
    QString path = QFileDialog::getOpenFileName(this, dlgTitle, curPath, filter);

    if (path.isEmpty())
        return;

    QFileInfo fileinfo(path);
    QDir::setCurrent(fileinfo.absolutePath());

    // 尝试打开一个Word文档，这里调用的是Documents对象的open方法
    mydocs->dynamicCall("Open(const QVariant&)",QVariant(path));

    // 获取当前word文档对象
    mydoc = myword->querySubObject("ActiveDocument");

    // 错误处理3
    if (mydoc == nullptr)
    {
        QMessageBox::information(this, QString("标题"), QString("获取当前文档失败"));
        return;
    }

    // 全选并获取文档内容，显示到文本框
    selection = mydoc->querySubObject("Range()");
    QString str = selection->property("Text").toString();
    ui->plainTextEdit->setPlainText(str); // 特别注意我这里用的是纯文本

    // 更改标题
    QString shortName = "test.docx";
    this->setWindowTitle(shortName);
    emit titleChanged(shortName);

    // 关闭文档、文档集合并退出
    mydoc->dynamicCall("Close()");
    delete mydoc;
    mydoc = NULL;

    myword->dynamicCall("Quit()");
    delete myword;
    myword = NULL;
}

// 保存word的action
void TFormDoc::on_actSaveWord_triggered()
{
    // 获取文本框内容
    QString str=ui->plainTextEdit->toPlainText();

    QString fileName="test.docx";
    QString path=QString(DocPath)+fileName;

//    QString curPath = QDir::currentPath();
//    QString dlgTitle = "保存一个word文档";
//    QString filter = "word 文档(*.docx);;word 文档(*.doc)";
//    QString path = QFileDialog::getOpenFileName(this, dlgTitle, curPath, filter);

//    if (path.isEmpty())
//        return;

//    QFileInfo fileinfo(path);
//    QDir::setCurrent(fileinfo.absolutePath());

    // 隐式打开一个word应用程序
    myword = new QAxObject("Word.Application");
    myword->setProperty("Visible", false);

    // 获取文档集合
    mydocs = myword->querySubObject("Documents");
    if (mydocs == nullptr)
    {
        QMessageBox::information(this, QString("标题"), QString("获取所有工作文档失败"));
        return;
    }

    // 创建一个word文档并获取当前激活的文档
    mydocs->dynamicCall("Add (void)");
    mydoc = myword->querySubObject("ActiveDocument");
    if (mydoc == nullptr)
    {
        QMessageBox::information(this, QString("标题"), QString("获取当前激活文档失败"));
        return;
    };

    // 写入文件内容
    selection = myword->querySubObject("Selection");
    selection->dynamicCall("TypeText(const QString&)",str);

    // 设置保存
    QVariant newFileName(path);
    QVariant fileFormat(16);
    QVariant LockComments(false);
    QVariant Password("");
    QVariant recent(true); // 如果为True，则将文档添加到“文件”菜单上最近使用的文件列表中。默认值为True
    QVariant writePassword("");
    QVariant ReadOnlyRecommended(false); // 如果为True，则无论何时打开文档，Word都将处于只读状态。默认值为False

    // 关闭文档并退出
    mydoc->dynamicCall("Close(boolean)", true);
    delete mydoc;
    mydoc = NULL;

    myword->dynamicCall("Quit(void)");
    delete myword;
    myword = NULL;
}

// 更改字体的action
void TFormDoc::on_actFont_triggered()
{
    QFont font=ui->plainTextEdit->font();

    bool ok;
    font=QFontDialog::getFont(&ok,font);
    ui->plainTextEdit->setFont(font);

}

