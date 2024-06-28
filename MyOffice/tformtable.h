#ifndef TFORMTABLE_H
#define TFORMTABLE_H

#include <QMainWindow>

#include    <QStandardItemModel>
#include    <QItemSelectionModel>
#include    <QLabel>

#include    "tdialogheaders.h"
#include    "tdialogsize.h"

QT_BEGIN_NAMESPACE
namespace Ui {class TFormTable;}
QT_END_NAMESPACE

class TFormTable : public QMainWindow
{
    Q_OBJECT

private:
    TDialogHeaders *dlgSetHeaders=nullptr; // 设置表头的对话框

    QStandardItemModel  *m_model;      // 数据模型
    QItemSelectionModel *m_selection;  // 选择模型

    // 状态栏组件
    QLabel  *labCellPos;    //当前单元格行列号
    QLabel  *labCellText;   //当前单元格内容

public:
    TFormTable(QWidget *parent = 0);
    ~TFormTable();

private slots:
    // 自定义槽函数，与QItemSelectionModel的currentChanged()信号连接
    void do_currentChanged(const QModelIndex &current, const QModelIndex &previous);

    // 表格操作槽函数
    void on_actSize_triggered();
    void on_actSetHeader_triggered();
    void on_actAppend_triggered();
    void on_actInsert_triggered();
    void on_actDelete_triggered();

    // 位置
    void on_actAlignLeft_triggered();
    void on_actAlignCenter_triggered();
    void on_actAlignRight_triggered();

    // 字体
    void on_actBold_triggered(bool checked);
    void on_actUnder_triggered(bool checked);
    void on_actItalic_triggered(bool checked);

    // 文件读写
    void on_actOpenExcel_triggered();
    void on_actSaveExcel_triggered();

private:
    Ui::TFormTable *ui;
};

#endif // TFORMTABLE_H
