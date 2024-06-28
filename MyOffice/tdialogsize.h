#ifndef TDIALOGSIZE_H
#define TDIALOGSIZE_H

#include <QDialog>

QT_BEGIN_NAMESPACE
namespace Ui {class TDialogSize;}
QT_END_NAMESPACE

// 表格大小设置
class TDialogSize : public QDialog
{
    Q_OBJECT

public:
    TDialogSize(QWidget *parent = 0);
    ~TDialogSize();

    int rowCount();
    int columnCount();
    void setRowColumn(int row, int column);

private slots:

private:
    Ui::TDialogSize *ui;
};

#endif // TDIALOGSIZE_H
