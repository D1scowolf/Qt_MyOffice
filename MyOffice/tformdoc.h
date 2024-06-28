#ifndef TFORMDOC_H
#define TFORMDOC_H

#include <QWidget>

#include <QAxWidget>
#include <QAxObject>
#include <QStatusBar>

QT_BEGIN_NAMESPACE
namespace Ui {class TFormDoc;}
QT_END_NAMESPACE

// 文档处理
class TFormDoc : public QWidget
{
    Q_OBJECT

public:
    TFormDoc(QWidget *parent = nullptr);
    ~TFormDoc();

    // 文本文件处理：.h .cpp .xml .json
    void loadFromFile(QString& aFileName);
    bool saveByIO_Whole(const QString& aFilename);

private slots:
    // 自定义统计字数
    void do_textChanged();

    // 文本文件处理的自定义action
    void on_actOpen_triggered();
    void on_actSave_triggered();
    // 更改字体的自定义action
    void on_actFont_triggered();
    // docx文件处理的自定义action：这里的docx需要是纯文本，不能带图片和表格
    void on_actOpenWord_triggered();
    void on_actSaveWord_triggered();

signals:
    // 检测tab标题的变化
    void titleChanged(QString title);

private:
    Ui::TFormDoc *ui;

    QAxObject *myword;      // Word应用程序指针
    QAxObject *mydocs;      // 文档集指针
    QAxObject *mydoc;       // 文档指针
    QAxObject *selection;   // Selection指针

    QStatusBar *statusBar;
};

#endif // TFORMDOC_H
